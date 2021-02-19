#!/usr/bin/env ruby
Dir.chdir("#{__dir__}/..")

exec "mvn compile exec:java -Dexec.mainClass=com.socioart.XlsxTemplate -Dexec.args='example/template.xlsx example/rendered.xlsx example/data.json'"
