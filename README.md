# XlsxTemplate

# Build

    # Executable .jar with all dependencies
    mvn package assembly:single

    # Run
    java -jar target/xlsx-template-jar-with-dependencies.jar compile template.xlsx rendered.xlsx data.json
    java -jar target/xlsx-template-jar-with-dependencies.jar list-pictures template.xls
    java -jar target/xlsx-template-jar-with-dependencies.jar replace-picture template.xlsx rendered.xlsx 'Sheet1!A1' image.png

    # Compile and execute locally (for development)
    mvn compile exec:java -Dexec.mainClass=com.socioart.XlsxTemplate -Dexec.args='compile template.xlsx rendered.xlsx data.json'

    # Render example to `example/rendered.xlsx` (requires Ruby)
    example/example.rb
