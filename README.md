# Omnidocx

Omnidocx is ruby gem that allows you to merge multiple docx (microsoft word) files into one, writing images to a docx file and making string replacements in the header, footer or main document content.
This gem works for docx files generated from microsoft word as well as google docs.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'omnidocx'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install omnidocx

## To Merge Documents

If you plan to have a header and footer applied to all the pages of the final document, then pass the document with header and footer as the first document in the array. Currently multiple different headers and footers are not supported.

While passing documents array make sure all the documents are from the same source, i.e., either microsoft word or google docs. Passing a mix of documents created from microsoft word and google docs might throw up namespace errors. 


```ruby
require 'omnidocx'

# To merge multiple docx files into one, you can use the following
# documents_to_merge is an array of documents (file paths) need to be merged and page_break is a boolean value if you want page breaks in b/w documents
$ Omnidocx::Docx.merge_documents(documents_to_merge=[], output_document_path, page_break)

# for e.g. if you had to merge two documents, just pass their entire paths in an array, if you need a page break in between documents then pass the page_break flag as true 
$ Omnidocx::Docx.merge_documents(['tmp/doc1.docx', 'tmp/doc2.docx'], 'tmp/output_doc.docx', true)
```

## To Write Images to a Document
```ruby
require 'omnidocx'

# To write images to a document, you can use the following
# images_to_write is an array of hashes, where each hash stores information about one image
$ Omnidocx::Docx.write_images_to_doc(images_to_write=[], input_document_path, output_document_path)

# Below is an example of the images_to_write array that you can pass in for images to be written to the doc
# image path, height and width are mandatory

    $ images_to_write = [ {
                          :path => "tmp/image1.jpg",     #URL || local path
                          :height => 500,
                          :width => 500,
                          :hdpi => 115,       #optional
                          :vdpi => 115        #optional
                          },
                          :path => "https://xyz.com/abc.jpeg",    #URL || local path
                          :height => 800,
                          :width => 500,
                          :hdpi => 115,       #optional
                          :vdpi => 115        #optional
                          }
                        ]

```

## For String Replacements

There are three different methods that can be used for string replacements

```ruby
require 'omnidocx'

# replacement_hash is a hash with keys present in the document that are to be replaced with their corresponding values

# For document content, you can use the following
$ Omnidocx::Docx.replace_doc_content(replacement_hash={}, input_document_path, output_document_path)

# For header content, you can use the following
$ Omnidocx::Docx.replace_header_content(replacement_hash={}, input_document_path, output_document_path)

# For footer content, you can use the following
$ Omnidocx::Docx.replace_footer_content(replacement_hash={}, input_document_path, output_document_path)

# Below is an example of how replacement_hash can be constructed 
$ replacement_hash = { "first_name" => "John", "last_name" => "Doe"}

```

Will be adding test specs soon.

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/parthnagori/omnidocx. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

