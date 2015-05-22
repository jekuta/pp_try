require 'zipruby'
require 'nokogiri'
require 'fileutils'

xml = Nokogiri::XML(Zip::Archive.open('test.pptx') { |z| z.fopen('ppt/slides/slide1.xml').read })

xml.xpath('//a:t').first.content = 'Changed again'

Zip::Archive.open('test.pptx', Zip::CREATE) { |z| z.add_or_replace_buffer('ppt/slides/slide1.xml', xml.to_s) }
