require "omnidocx/version"
require 'nokogiri'
require 'zip'
require 'tempfile'
require 'mime/types'
require 'open-uri'


module Omnidocx
  class Docx
    DOCUMENT_FILE_PATH = 'word/document.xml'
    RELATIONSHIP_FILE_PATH = 'word/_rels/document.xml.rels'
    CONTENT_TYPES_FILE = '[Content_Types].xml'
    HEADER_RELS_FILE_PATH = 'word/_rels/header1.xml.rels'
    FOOTER_RELS_FILE_PATH = 'word/_rels/footer1.xml.rels'
    STYLES_FILE_PATH = "word/styles.xml"
    HEADER_FILE_PATH = "word/header1.xml"
    FOOTER_FILE_PATH = "word/footer1.xml"

    MEDIA_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    EMUSPERINCH = 914400
    EMUSPERCM = 360000
    HORIZONTAL_DPI = 115
    VERTICAL_DPI = 117

    NAMESPACES = {
      "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }

    IMAGE_ELEMENT = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:rsidR="00F127EA" w:rsidRDefault="00F127EA" w:rsidP="00BF4C96"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/><w:lang w:eastAsia="en-IN"/></w:rPr><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="" cy=""/><wp:effectExtent l="0" t="0" r="2540" b="1905"/><wp:docPr id="" name=""/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed=""><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="" cy=""/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'


    def self.write_images_to_doc(images_to_write=[], doc_path, final_path)

      temp_file = Tempfile.new('docxedit-')

      #every docx file is ultimately a zip file with the extension as docx
      @document_zip = Zip::File.new(doc_path)
      #reading the document content xml from the zip
      @document_content = @document_zip.read(DOCUMENT_FILE_PATH)
      @document_xml = Nokogiri::XML @document_content

      #every docx file has one body tag which essentially contains all the content of the doc
      @body = @document_xml.xpath("//w:body")

      @rel_doc = ""
      @cont_type_doc = ""

      cnt = 20
      media_hash = {}

      #to maintain a list of all the content type info to be added upon adding media with different extensions
      media_content_type_hash = {}


      @document_zip.entries.each do |e|
        if e.name == RELATIONSHIP_FILE_PATH
          in_stream = e.get_input_stream.read
          @rel_doc = Nokogiri::XML in_stream    #Relationships XML
        end
        if e.name == CONTENT_TYPES_FILE
          in_stream = e.get_input_stream.read
          @cont_type_doc = Nokogiri::XML in_stream  #Content types XML to be updated later on with the additional media type info
        end
      end

      Zip::OutputStream.open(temp_file.path) do |zos|

        @document_zip.entries.each do |e|
          unless [DOCUMENT_FILE_PATH, RELATIONSHIP_FILE_PATH, CONTENT_TYPES_FILE].include?(e.name)
            #writing the files not needed to be edited back to the new zip
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        images_to_write.each_with_index do |img, index|
          data = ''

          #checking if image path is a url or a local path
          uri = URI.parse(img[:path])
          if %w( http https ).include?(uri.scheme)
            data = Kernel.open(img[:path]).read rescue nil
          else
            File.open(img[:path], 'rb') do |f|
              data = f.read rescue nil
            end
          end

          #if image path is readable
          if !data.empty?
            img_url_no_params = img[:path].gsub(/\?.*/,'')
            extension = File.extname(img_url_no_params).split(".").last

            if !media_content_type_hash.keys.include?(extension.split(".").last)
              #making an entry for a new media type
              media_content_type_hash["#{extension}"] = MIME::Types.type_for(img_url_no_params)[0].to_s
            end

            zos.put_next_entry("word/media/image#{cnt}.#{extension}")
            zos.print data     #storing the image in the new zip

            new_rel_node = Nokogiri::XML::Node.new("Relationship", @rel_doc)
            new_rel_node["Id"] = "rid#{cnt}"
            new_rel_node["Type"] = MEDIA_TYPE
            new_rel_node["Target"] = "media/image#{cnt}.#{extension}"
            @rel_doc.at('Relationships').add_child(new_rel_node)      #adding a new relationship node to the relationships xml

            hdpi = img[:hdpi] || HORIZONTAL_DPI
            vdpi = img[:vdpi] || VERTICAL_DPI

            #calculating the width and height of the image in EMUs, the format accepted by docx files
            widthEmus = (img[:width].to_i / hdpi.to_i * EMUSPERINCH)
            heightEmus = (img[:height].to_i / vdpi.to_i * EMUSPERINCH)

            #creating a new drawing element with info like rid, height, width,etc.
            @image_element_xml = Nokogiri::XML IMAGE_ELEMENT
            @image_element_xml.xpath("//w:drawing", NAMESPACES).each do |dr_node|
              docPr = dr_node.xpath(".//wp:docPr", NAMESPACES).last
              docPr["name"] = "image#{cnt}.#{extension}"
              docPr["id"] = "#{cnt}"

              extent = dr_node.xpath(".//wp:extent", NAMESPACES).last
              extent["cx"] = widthEmus.to_s
              extent["cy"] = heightEmus.to_s

              ext = dr_node.xpath(".//a:ext", NAMESPACES).last
              ext["cx"] = widthEmus.to_s
              ext["cy"] = heightEmus.to_s

              pic_cNvPr = dr_node.xpath(".//pic:cNvPr", NAMESPACES).last
              pic_cNvPr["name"] = "image#{cnt}.#{extension}"
              pic_cNvPr["id"] = "#{cnt}"

              blip = dr_node.xpath(".//a:blip", NAMESPACES).last
              blip.attributes["embed"].value = "rid#{cnt}"
            end

            #appending the drawing element to the document's body
            @body.children.last.add_previous_sibling(@image_element_xml.xpath("//w:p").last.to_xml)

            media_hash[cnt] = index
          end
          cnt+=1
        end

        #updating the content type info
        media_content_type_hash.each do |ext, cont_type|
          new_default_node = Nokogiri::XML::Node.new("Default", @cont_type_doc)
          new_default_node["Extension"] = ext
          new_default_node["ContentType"] = cont_type
          @cont_type_doc.at("Types").add_child(new_default_node)
        end

        #writing the content types xml to the new zip
        zos.put_next_entry CONTENT_TYPES_FILE
        zos.print @cont_type_doc.to_xml

        #writing the relationships xml to the new zip
        zos.put_next_entry RELATIONSHIP_FILE_PATH
        zos.print @rel_doc.to_xml

        #writing the updated document content xml to the new zip
        zos.put_next_entry DOCUMENT_FILE_PATH
        zos.print @document_xml.to_xml
      end

      #moving the temporary docx file to the final_path specified by the user
      FileUtils.mv(temp_file.path, final_path)
    end

    def self.merge_documents(documents_to_merge=[], final_path, page_break)
      temp_file = Tempfile.new('docxedit-')

      if documents_to_merge.count < 2
        return "Pass atleast two documents to be merged"   #minimum two documents required to merge
      end

      #first document to which the others will be appended (header/footer will be picked from this document)
      @main_document_zip = Zip::File.new(documents_to_merge.first)
      @main_document_content = @main_document_zip.read(DOCUMENT_FILE_PATH)
      @main_document_xml = Nokogiri::XML @main_document_content
      @main_body = @main_document_xml.xpath("//w:body")
      @rel_nodes = ""
      @rel_doc = ""
      @cont_type_doc = ""
      @style_doc = ""
      doc_cnt = 0
      #cnt variable to construct relationship ids, taken a high value 100 to avoid duplication
      cnt = 100
      tbl_cnt = 10
      #hash to store information about the media files and their corresponding new names
      media_hash = {}
      #rid_hash to store relationship information
      rid_hash = {}
      #table hash to store information if any tables present
      table_hash = {}
      #head_foot_media hash to store if any media files present in header/footer
      head_foot_media = {}
      #a counter for docPr element in the main document body
      docPr_id = 100

      #array to store content type information about media extensions
      default_extensions = []
      #array to store override content type information
      override_partnames = []

      #array to store information about additional content types other than the ones present in the first(main) document
      additional_cont_type_entries = []


      @main_document_zip.entries.each do |e|
        if e.name == RELATIONSHIP_FILE_PATH
          @in_stream = e.get_input_stream.read
          @rel_doc = Nokogiri::XML @in_stream        #Relationship XML
          @rel_nodes = @rel_doc.css "Relationship"
        end
        if e.name == CONTENT_TYPES_FILE
          in_stream = e.get_input_stream.read
          @cont_type_doc = Nokogiri::XML in_stream       #Content types XML to be updated later on with the additional media type info
          default_nodes = @cont_type_doc.css "Default"
          override_nodes = @cont_type_doc.css "Override"
          default_nodes.each do |node|
            default_extensions << node["Extension"]
          end
          override_nodes.each do |node|
            override_partnames << node["PartName"]
          end
        end
        if e.name == STYLES_FILE_PATH
          in_stream = e.get_input_stream.read
          @style_doc = Nokogiri::XML in_stream      #Styles XML to be updated later on with the additional tables info
        end
      end


      #opening a new zip for the final document
      Zip::OutputStream.open(temp_file.path) do |zos|
        documents_to_merge.each do |doc_path|
          media_hash["doc#{doc_cnt}"] = {}
          rid_hash["doc#{doc_cnt}"] = {}
          head_foot_media["doc#{doc_cnt}"] = []
          table_hash["doc#{doc_cnt}"] = {}
          zip_file = Zip::File.new(doc_path)

          zip_file.entries.each do |e|
            if [HEADER_RELS_FILE_PATH, FOOTER_RELS_FILE_PATH].include?(e.name)
              hf_content = e.get_input_stream.read
              hf_xml = Nokogiri::XML hf_content
              hf_xml.css("Relationship").each do |rel_node|
                #media file names in header & footer need not be changed as they will be picked from the first document only and not the subsequent documents, so no chance of duplication
                head_foot_media["doc#{doc_cnt}"] << rel_node["Target"].gsub("media/","")
              end
            end
            if e.name == CONTENT_TYPES_FILE
              cont_types = e.get_input_stream.read
              cont_type_xml = Nokogiri::XML cont_types
              default_nodes = cont_type_xml.css "Default"
              override_nodes = cont_type_xml.css "Override"

              default_nodes.each do |node|
                #checking if extension type already present in the content types xml extracted from the first document
                if !default_extensions.include?(node["Extension"]) && !node.to_xml.empty?
                  additional_cont_type_entries << node
                  default_extensions << node["Extension"]    #extra extension type to be added to the content types XML
                end
              end

              override_nodes.each do |node|
                #checking if override content tpye infoalready present in the content types xml extracted from the first document
                if !override_partnames.include?(node["PartName"]) && !node.to_xml.empty?
                  additional_cont_type_entries << node
                  override_partnames << node["Partname"]       #extra content type info to be added to the content types XML
                end
              end
            end
          end


          zip_file.entries.each do |e|
            unless e.name == DOCUMENT_FILE_PATH || [RELATIONSHIP_FILE_PATH, CONTENT_TYPES_FILE, STYLES_FILE_PATH].include?(e.name)
              if e.name.include?("word/media/image")
                if !head_foot_media["doc#{doc_cnt}"].include?(e.name.gsub("word/media/",""))
                  #renaming media files with a higher counter to avoid duplicaiton in case multiple documents have images present
                  e_name = e.name.gsub(/image[0-9]*./,"image#{cnt}.")
                  #writing the media file back to the new zip with the new name
                  zos.put_next_entry(e_name)
                  zos.print e.get_input_stream.read
                  #storing the old media file name to new media file name to mapping in the media hash
                  media_hash["doc#{doc_cnt}"][e.name.gsub("word/media/","")] = cnt
                  cnt+=1
                else
                  #writing the media files present in the header and footer as their names are not needed to be changed
                  zos.put_next_entry(e.name)
                  zos.print e.get_input_stream.read
                end
              else
                #writing the files not needed to be edited back to the new zip (only from the first document, so as to avoid duplication)
                if doc_cnt == 0
                  zos.put_next_entry(e.name)
                  zos.print e.get_input_stream.read
                end
              end
            end
          end

          if doc_cnt == 0
            doc_content = @main_document_xml      #first document's content XML
          else
            document_content = zip_file.read(DOCUMENT_FILE_PATH)
            doc_content = Nokogiri::XML document_content      #subsequent documents' content XML
          end

          #updating the stlye ids in the table elements present in the document content XML
          doc_content.xpath("//w:tbl").each do |tbl_node|
            tblStyle = tbl_node.xpath('.//w:tblStyle').last

            table_hash["doc#{doc_cnt}"]["#{tblStyle.attributes['val'].value}"] = tbl_cnt
            tblStyle.attributes['val'].value = tblStyle.attributes['val'].value.gsub(/[0-9]+/,"#{tbl_cnt}")
            tbl_cnt+=1
          end

          #updating the relationship ids with the new media file names in the relationships XML
          if doc_cnt == 0
            zip_file.entries.each do |e|
              if e.name == RELATIONSHIP_FILE_PATH
                @rel_nodes.each do |node|
                  if node.values.to_s.include?("image")
                    i = media_hash["doc#{doc_cnt}"]["#{node['Target']}".gsub("media/","")]
                    target_val = node["Target"].gsub(/image[0-9]*./,"image#{i}.")
                    node["Target"] = target_val
                    rid_hash["doc#{doc_cnt}"]["#{node['Id']}"] = "#{i}"
                    node.attributes["Id"].value = node.attributes["Id"].value.gsub(/[0-9]+/,"#{i}")
                  end
                end
              end
              #adding the table style information to the styles xml, if any tables present in the document being merged
              if e.name == STYLES_FILE_PATH
                table_nodes = @style_doc.xpath('//w:style').select{|n| n.attributes["type"].value == "table"}
                table_nodes.each do |table_node|
                  tab_val = table_hash["doc#{doc_cnt}"]["#{table_node.attributes['styleId'].value}"]
                  table_node.attributes['styleId'].value = table_node.attributes['styleId'].value.gsub(/[0-9]+/,"#{tab_val}")
                end
              end
            end
          else
            zip_file.entries.each do |e|
              if e.name == RELATIONSHIP_FILE_PATH
                input_stream = e.get_input_stream.read
                rel_xml = Nokogiri::XML input_stream
                rel_xml_nodes = rel_xml.css "Relationship"
                rel_xml_nodes.each do |node|
                  if node.values.to_s.include?("image")
                    i = media_hash["doc#{doc_cnt}"]["#{node['Target']}".gsub("media/","")]
                    target_val = node["Target"].gsub(/image[0-9]*./,"image#{i}.")
                    rid_hash["doc#{doc_cnt}"]["#{node['Id']}"] = "#{i}"

                    new_rel_node = Nokogiri::XML::Node.new("Relationship", @rel_doc)
                    new_rel_node["Id"] = node.attributes["Id"].value.gsub(/[0-9]+/,"#{i}")
                    new_rel_node["Type"] = node["Type"]
                    new_rel_node["Target"] = target_val

                    #adding the extra relationship nodes for the media files from the subsequent documents (apart from first) to the relationship XML
                    @rel_doc.at('Relationships').add_child(new_rel_node)
                  end
                end
              end

              if e.name == STYLES_FILE_PATH
                input_stream = e.get_input_stream.read
                style_xml = Nokogiri::XML input_stream
                table_nodes = style_xml.xpath("//w:style").select{|n| n.attributes["type"].value == "table" && n.attributes["styleId"].value != "TableNormal"}
                table_nodes.each do |table_node|
                  tab_val = table_hash["doc#{doc_cnt}"]["#{table_node.attributes['styleId'].value}"]
                  table_node.attributes['styleId'].value = table_node.attributes['styleId'].value.gsub(/[0-9]+/,"#{tab_val}")
                  #adding extra table style nodes to the styles xml, if any tables present in the document being merged
                  @style_doc.xpath("//w:styles").children.last.add_next_sibling(table_node.to_xml)
                end
              end
            end
          end

          #updting the id and rid values for every drawing element in the document XML with the new counters
          doc_content.xpath("//w:drawing").each do |dr_node|
            blip = dr_node.xpath(".//a:blip", NAMESPACES).last
            next if blip.nil?
            i = rid_hash["doc#{doc_cnt}"][blip.attributes["embed"].value]
            blip.attributes["embed"].value = blip.attributes["embed"].value.gsub(/[0-9]+/,i)
            docPr = dr_node.xpath(".//wp:docPr").last
            docPr["id"] = #{docPr_id}
            docPr_id+=1
          end


          if doc_cnt > 0
            w_p_nodes = doc_content.xpath("//w:p")
            #pulling out the <w:p> elements fromt the document body to be appended to the main document's body
            body_nodes = doc_content.xpath('//w:body').children[0..doc_content.xpath('//w:body').children.count-2]

            #adding a page break between documents being merged
            if doc_cnt > 1 && page_break
              @main_body.children.last.add_previous_sibling('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
            end
            #appending the body_nodes to main document's body
            @main_body.children.last.add_previous_sibling(body_nodes.to_xml)
          end

          doc_cnt+=1
        end

        #writing the updated styles XML to the new zip
        zos.put_next_entry(STYLES_FILE_PATH)
        zos.print @style_doc.to_xml

        #writing the updated relationships XML to the new zip
        zos.put_next_entry(RELATIONSHIP_FILE_PATH)
        zos.print @rel_doc.to_xml

        zos.put_next_entry(CONTENT_TYPES_FILE)
        additional_cont_type_entries.each do |node|
          #adding addtional content type nodes to the content type XML
          @cont_type_doc.at("Types").add_child(node)
        end
        #writing the updated content types XML to the new zip
        zos.print @cont_type_doc.to_xml

        #writing the updated document content XML to the new zip
        zos.put_next_entry(DOCUMENT_FILE_PATH)
        zos.print @main_document_xml.to_xml
      end

      #moving the temporary docx file to the final_path specified by the user
      FileUtils.mv(temp_file.path, final_path)
    end

    def self.replace_doc_content(replacement_hash={}, template_path, final_path)
      @template_zip = Zip::File.new(template_path)
      @template_content = @template_zip.read(DOCUMENT_FILE_PATH)

      #replacing the keys with values in the document content xml
      replacement_hash.each do |key,value|
        @template_content.force_encoding("UTF-8").gsub!(key,value)
      end

      temp_file = Tempfile.new('docxedit-')

      Zip::OutputStream.open(temp_file.path) do |zos|

        @template_zip.entries.each do |e|
          unless e.name == DOCUMENT_FILE_PATH
            #writing the files not needed to be edited back to the new zip
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        #writing the updated document content xml to the new zip
        zos.put_next_entry DOCUMENT_FILE_PATH
        zos.print @template_content
      end

      #moving the temporary docx file to the final_path specified by the user
      FileUtils.mv(temp_file.path, final_path)
    end

    def self.replace_header_content(replacement_hash={}, template_path, final_path)
      @template_zip = Zip::File.new(template_path)

      @header_content = ''
      @template_zip.entries.each do |e|
        if e.name == HEADER_FILE_PATH
          @header_content = e.get_input_stream.read
        end
      end

      replacement_hash.each do |key,value|
        @header_content.force_encoding("UTF-8").gsub!(key,value)
      end

      temp_file = Tempfile.new('docxedit-')

      Zip::OutputStream.open(temp_file.path) do |zos|

        @template_zip.entries.each do |e|
          unless e.name == HEADER_FILE_PATH
            #writing the files not needed to be edited back to the new zip
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        #writing the updated document content xml to the new zip
        zos.put_next_entry HEADER_FILE_PATH
        zos.print @header_content
      end

      #moving the temporary docx file to the final_path specified by the user
      FileUtils.mv(temp_file.path, final_path)
    end

    def self.replace_footer_content(replacement_hash={}, template_path, final_path)
      @template_zip = Zip::File.new(template_path)

      @footer_content = ''
      @template_zip.entries.each do |e|
        if e.name == FOOTER_FILE_PATH
          @footer_content = e.get_input_stream.read
        end
      end

      replacement_hash.each do |key,value|
        @footer_content.force_encoding("UTF-8").gsub!(key,value)
      end

      temp_file = Tempfile.new('docxedit-')

      Zip::OutputStream.open(temp_file.path) do |zos|

        @template_zip.entries.each do |e|
          unless e.name == FOOTER_FILE_PATH
            #writing the files not needed to be edited back to the new zip
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        #writing the updated document content xml to the new zip
        zos.put_next_entry FOOTER_FILE_PATH
        zos.print @footer_content
      end

      #moving the temporary docx file to the final_path specified by the user
      FileUtils.mv(temp_file.path, final_path)
    end

  end
end
