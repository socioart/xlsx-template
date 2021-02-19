#!/usr/bin/env ruby

require "open3"
def sh(*args)
  o, e, s = Open3.capture3(*args, out: $stderr)
  unless s.success?
    $stderr.puts e
    exit s.to_i
  end
  o
end

Dir.chdir("#{__dir__}/..")


sh "mvn compile"
puts sh "mvn exec:java -Dexec.mainClass=com.socioart.XlsxTemplate -Dexec.args='list-pictures example/template.xlsx'"
sh "mvn exec:java -Dexec.mainClass=com.socioart.XlsxTemplate -Dexec.args='replace-picture example/template.xlsx example/rendered.xlsx Sheet1!G2 example/replace.png'"
sh "mvn exec:java -Dexec.mainClass=com.socioart.XlsxTemplate -Dexec.args='compile example/rendered.xlsx example/rendered.xlsx example/data.json'"
