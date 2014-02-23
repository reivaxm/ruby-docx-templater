require 'nokogiri'

module DocxTemplater
  class TemplateProcessor
    include Logging
    attr_reader :data, :escape_html

    # data is expected to be a hash of symbols => string or arrays of hashes.
    def initialize(data, escape_html = true)
      @data = data
      @escape_html = escape_html
    end

    def render(document)
      document.force_encoding(Encoding::UTF_8) if document.respond_to?(:force_encoding)
      data.each do |key, value|
        if value.class == Array
          document = enter_multiple_values(document, key)
          document.gsub!("#SUM:#{key.to_s.upcase}#", value.count.to_s)
        else
          document.gsub!("$#{key.to_s.upcase}$", safe(value))
        end
      end
      document
    end

    # Proceed Document update
    # @params document [Nokogiri::XML] Document XML to be update
    # @return document [Nokogiri::XML] DOcument updated
    def update_document(document)
      document.force_encoding(Encoding::UTF_8) if document.respond_to?(:force_encoding)
      xml = Nokogiri::XML(document)
      xml = substitute_element(xml, data)
      xml.to_s
    end
    private

    def safe(text)
      if escape_html
        text.to_s.gsub('&', '&amp;').gsub('>', '&gt;').gsub('<', '&lt;')
      else
        text.to_s
      end
    end


    def enter_multiple_values(document, key)
      logger.debug("enter_multiple_values for: #{key}")
      # TODO: ideally we would not re-parse xml doc every time
      xml = Nokogiri::XML(document)

      begin_row = "#BEGIN_ROW:#{key.to_s.upcase}#"
      end_row = "#END_ROW:#{key.to_s.upcase}#"
      begin_row_template = xml.xpath("//w:tr[contains(., '#{begin_row}')]", xml.root.namespaces).first
      end_row_template = xml.xpath("//w:tr[contains(., '#{end_row}')]", xml.root.namespaces).first
      logger.debug("begin_row_template: #{begin_row_template.to_s}")
      logger.debug("end_row_template: #{end_row_template.to_s}")
      fail "unmatched template markers: #{begin_row} nil: #{begin_row_template.nil?}, #{end_row} nil: #{end_row_template.nil?}. This could be because word broke up tags with it's own xml entries. See README." unless begin_row_template && end_row_template

      row_templates = []
      row = begin_row_template.next_sibling
      while row != end_row_template
        row_templates.unshift(row)
        row = row.next_sibling
      end
      logger.debug("row_templates: (#{row_templates.count}) #{row_templates.map(&:to_s).inspect}")

      # for each data, reversed so they come out in the right order
      data[key].reverse.each do |each_data|
        logger.debug("each_data: #{each_data.inspect}")

        # dup so we have new nodes to append
        row_templates.map(&:dup).each do |new_row|
          logger.debug("   new_row: #{new_row}")
          innards = new_row.inner_html
          matches = innards.scan(/\$EACH:([^\$]+)\$/)
          unless matches.empty?
            logger.debug("   matches: #{matches.inspect}")
            matches.map(&:first).each do |each_key|
              logger.debug("      each_key: #{each_key}")
              innards.gsub!("$EACH:#{each_key}$", safe(each_data[each_key.downcase.to_sym]))
            end
          end
          # change all the internals of the new node, even if we did not template
          new_row.inner_html = innards
          # DocxTemplater::log("new_row new innards: #{new_row.inner_html}")

          begin_row_template.add_next_sibling(new_row)
        end
      end
      (row_templates + [begin_row_template, end_row_template]).map(&:unlink)
      xml.to_s
    end

    # Substitute element in document
    # @param document [XML] Nokogiri document from docx
    # @param data [Hash] Data for exchange
    # @param tag_loop [Boolean] Set true if substitue ina loop (default : false)
    # @return [XML] return altered document
    def substitute_element(document, data, tag_loop=nil)
      data.each do |k,v|
        if v.is_a? Array
            logger.debug("Multiple value for key : #{k}")
            document = update_xml_between_tag(document,k) do |tag,xml|
              out = nil
              # On ech value iterate data
              v.each do |d|
                # Duplicate original extract
                origin_node = xml.dup
                # Treat contenat
                out_loop = substitute_element(origin_node,d,tag)
                # And creplace or concat
                if out.nil?
                  out = out_loop
                else
                  out = Nokogiri::XML::DocumentFragment.parse(out.to_xml + out_loop.to_xml)
                end
              end
              out
            end
            document = replace_value(document,"SUM:#{k.to_s.upcase}",v.count)
        else
          logger.debug("Substitute : #{k}")
          unless tag_loop.nil?
            document = replace_value(document,"EACH_#{tag_loop.to_s.upcase}:#{k.to_s.upcase}",v)
          else
            document = replace_value(document,k,v)
          end
        end
      end
      return document
    end

    # This function replace tag in document by the value in the HASH
    # @param document [Nokogiri::XML] Document
    # @param key [Symbole] Tag in the document
    # @param value [Strinng] Data to be replace
    # @return xml [Nokogiri::XML] return altered document
    def replace_value(xml,key,value)
      key = key.to_s.upcase
      logger.debug("Trying replace #{key} by #{value}")
      # Try getting xml namespace
      begin
        ns = xml.root.namespace
      rescue NoMethodError
        ns = nil
      end
      # Defined Document prefix from namespace
      prefix = ''
      unless ns.nil?
        prefix = "#{ns.prefix}:"
      end
      node = xml.xpath(".//#{prefix}t[contains(.,'$#{key}$')]")
      logger.debug("Replacement node : #{node.inspect}")
      unless node.count == 0
        # Escape html if needed
        if escape_html
          content = value.to_s.gsub('&', '&amp;').gsub('>', '&gt;').gsub('<', '&lt;')
        else
          content = value.to_s
        end
        # And replace in node
        node.each{|n| n.content = n.content.sub(/\$#{key}\$/,content)}
      end
      return xml
    end

    # This function alter xml node between tag
    # Give 2 elment in a block, tag and extracted nodes
    # elements between tag for edit befor return xml document
    # @param xml [Nokogiri::XML] Nokogiri xml dcument
    # @param tag [Symbole] Tag in document for extract content
    # @return [Nokogiri::XML] return the document altered
    def update_xml_between_tag(xml,tag)
      # Try getting xml namespace
      begin
        ns = xml.root.namespace
      rescue NoMethodError
        ns = nil
      end
      # Defined Document prefix from namespace
      prefix = ''
      unless ns.nil?
        prefix = "#{ns.prefix}:"
      end
      # Search XML node who contain table or String marker
      begin_search = "[contains(.,'#BEGIN_ROW:#{tag.to_s.upcase}#')]"
      #logger.debug("Xpath Search : #{begin_search}")
      # Try talbe
      begin_row_template = xml.xpath(".//#{prefix}tr#{begin_search}",ns).first
      #logger.debug("Prefix tr : #{begin_row_template.inspect}")
      # Try line
      begin_row_template ||= xml.xpath(".//#{prefix}p#{begin_search}",ns).first
      #logger.debug("Prefix p : #{begin_row_template.inspect}")
      # Exit if marker not found
      if begin_row_template.nil?
        logger.debug "Unmatched template marker : #BEGIN_ROW:#{tag.to_s.upcase}#, abort this marker"
        return xml
      end
      # Detect correspondant tag
      tag_name = begin_row_template.name
      end_search = "[contains(.,'#END_ROW:#{tag.to_s.upcase}#')]"
      #logger.debug("Xpath Search : #{end_search}")
      # Search end tag marker
      end_row_template = xml.xpath(".//#{prefix}#{tag_name}#{end_search}",ns).first
      #logger.debug("Prefix #{tag_name} : #{begin_row_template.inspect}")
      # Exit if not found
      if begin_row_template.nil?
        logger.debug "Unmatched template marker : #END_ROW:#{tag.to_s.upcase}#, abort this marker"
        return xml
      end

      # Extract node in array
      row_templates = []
      row = begin_row_template.next_sibling
      while row != end_row_template
        row_templates << row
        row = row.next_sibling
      end
      # Remove profix if exist
      content_string = row_templates.inject(""){|s,n| s+= n.to_xml}
      content_string.gsub!(prefix,'') unless ns.nil?
      # Generate nex xml
      content_xml = Nokogiri::XML::DocumentFragment.parse(content_string)

      # Execute some action on XML node
      content_update = yield(tag,content_xml)
      # Add updatre content
      begin_row_template.add_next_sibling(content_update)
      # Remove all unused node
      (row_templates + [begin_row_template, end_row_template]).map(&:unlink)
      return xml
    end
  end
end
