module Docx
  module TemplateHandler

    START = 'REPEAT:'
    FINISH = 'UNREPEAT:'
    EACH = 'EACH:'
    PARAGRAPH_XPATH = '//w:p'
    TEXT_NODE_XPATH = './/w:t'
    BUX = '$'
    COLON = ':'

    PR = 'Pr'
    P = 'p'
    W = 'w'

    # Now supported only one-depth cycles and cycle begin/end must be in the separate rows
    def self.insert(hash, doc)
      paragraphs = doc.xpath(PARAGRAPH_XPATH)
      paragraphs.each do |paragraph|
        text_nodes = paragraph.xpath(TEXT_NODE_XPATH)
        text_nodes.each do |text_node|
          value = text_node.content
          next unless value[0] == BUX && value[-1] == BUX
          key = value[1..-2]
          if key.include? COLON
            if key.start_with? START
              key.sub! START, FINISH
              fin_value = value.sub START, FINISH
              paragraph, repeat = bound_repeatable_piece(text_node, fin_value)
              elems_for_repeating = key.sub FINISH, ''
              if hash[elems_for_repeating] && (hash[elems_for_repeating].kind_of? Enumerable)
                hash[elems_for_repeating].each do |repeated_elem|
                  repeat.each do |repeatable_paragraph|
                    dup_par = repeatable_paragraph.dup
                    handle_paragraph dup_par, repeated_elem
                    paragraph.add_next_sibling dup_par
                    paragraph = dup_par
                  end
                end
              end
              repeat.each do |par|
                par.remove
              end
              break
            else
              vars = key.split COLON
              obj = hash[vars[0]]
              text_node.content = if obj.is_a? Hash
                                    obj[vars[1]]
                                  else
                                    begin
                                      obj.send vars[1]
                                    rescue NoMethodError => e
                                      value
                                    end
                                  end
            end
          else
            text_node.content = hash[key] || value
          end
        end
      end
    end

    private


    def self.handle_paragraph(paragraph, hash)
      paragraph.xpath(TEXT_NODE_XPATH).each do |text_node|
        value = text_node.content
        next unless value[0] == BUX && value[-1] == BUX
        key = value[1..-2]
        if key.include? COLON
          if key.start_with? EACH
            key.sub! EACH, ''
            text_node.content = hash[key] || value
          end
        end
      end
    end

    def self.bound_repeatable_piece(node, fin_value)
      # Tag name consists from prefix (usually, 'w') and name (i.e. 'r', 'rPr' and so on)
      temp_key = paragraph_containing node
      paragraph = temp_key.next_sibling
      temp_key.remove
      #next unless paragraph
      repeat = []
      while true
        result = find_finish paragraph, fin_value
        #if result.kind_of? Array
        #  new_children = Nokogiri::XML::NodeSet.new paragraph.document
        #  result.each do |node|
        #    new_children.push node
        #  end
        #  paragraph.children = new_children
        #  repeat.push paragraph
        #  break
        #end
        unless result
          # Move 'paragraph' back and remove ending delimiter
          temp_key = paragraph
          paragraph = paragraph.previous_sibling
          temp_key.remove
          break
        end
        #paragraph.xpath('.//w:t')
        repeat.push result
        paragraph = paragraph.next_sibling
        break unless paragraph
      end
      return paragraph, repeat
    end

    def self.find_finish(node, finish_str)

      # Now think that repeating will be bounded by paragraphs

      return node if node.name.end_with? PR
      children = []
      node.children.each do |child|
        ## Be careful, result = nil in the case, when we found finish_str
        #result = find_finish child, finish_str
        #unless result                                                             # '$FINISH:...$' returns nil
        #  return nil if children.empty?                                           # <w:t>$...$</w:t> returns nil here
        #  return nil if children.size == 1 && node.name+PR == children.first.name # <w:p><w:pPr>...</w:pPr><w:t>...</w:t> returns nil here
        #  return children
        #else
        #  # If node is returned, then... If array, then...
        #  if result.kind_of? Nokogiri::XML::Node
        #    children.push child
        #  elsif result.kind_of? Array
        #    new_children = Nokogiri::XML::NodeSet.new child.document
        #    result.each do |res_node|
        #      new_children.push res_node
        #    end
        #    child.children = new_children
        #    children.push child
        #    return children
        #  end
        #end
      end

      # WARNING: It seems, situation exists, when finish_str willn't contain in the children but will be in the node
      if node.content == finish_str
        return nil
      else
        return node
      end
    end


    def self.paragraph_containing node
      return nil unless node.namespace.prefix == W
      return node if node.name == P
      while (node = node.parent).name != P
      end
      node
    end
  end
end