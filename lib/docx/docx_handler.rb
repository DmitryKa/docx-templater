require 'zip'
require 'nokogiri'
require 'docx/template_handler'

module Docx
  class DocxHandler
    include TemplateHandler

    DOC_XML_ARCHIVE_PATH = 'word/document.xml'

    def self.open(path, &block)
      self.new(path, &block)
    end

    def initialize(path, &block)
      @replace = {}
      if block_given?
        @zip = Zip::File.open(path)
        yield(self)
        @zip.close
      else
        @zip = Zip::File.open(path)
      end
    end

    def insert(arg_hash)
      raise TypeError unless arg_hash.kind_of? Hash
      xml = @zip.read(DOC_XML_ARCHIVE_PATH)
      into_doc = Nokogiri::XML(xml) { |x| x.noent } #TODO what is noent?
      TemplateHandler::insert arg_hash, into_doc
      @replace[DOC_XML_ARCHIVE_PATH] = into_doc.serialize save_with: 0
    end

    def save(path)
      Zip::File.open(path, Zip::File::CREATE) do |out|
        @zip.each do |entry|
          out.get_output_stream(entry.name) do |o|
            if @replace[entry.name]
              o.write(@replace[entry.name])
            else
              o.write(@zip.read(entry.name))
            end
          end
        end
      end
      @zip.close
    end
  end
end