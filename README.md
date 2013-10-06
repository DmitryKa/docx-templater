# DocxTemplater

PURE RUBY GEM (works on Windows)! I tried use ready solutions, but some of them hasn't neccessary functions, other one didn't work in Win.

## Installation

Add this line to your application's Gemfile:

    gem 'docx-templater2'

And then execute:

    $ bundle install

Or install it yourself as:

    $ gem install docx_templater

## Usage

Can work with:
- $SOME_VALUE$ : replace template stub (if value for replacing will be provided);
- $REPEAT:SOME_COLLECTION$
 (at least, one new row)
 $EACH:SOME_VALUE_OF_COLLECTION_ITEM$
 (at least, one new row)
 $UNREPEAT:SOME_COLLECTION$;

Repeating commands must be at the separate paragraphs.
Note, that created template may look good but have bad structure for templater, so you should open docx as a zip archive, open "word/document.xml" and edit it

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
