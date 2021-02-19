# XlsxTemplate

# Build

    # Executable .jar with all dependencies
    mvn package assembly:single
    java -jar target/xlsx-template-jar-with-dependencies.jar template.xlsx rendered.xlsx data.json

    # Compile and execute locally (for development)
    mvn compile exec:java -Dexec.mainClass=com.socioart.XlsxTemplate -Dexec.args='template.xlsx rendered.xlsx data.json'

    # Render example to `example/rendered.xlsx` (requires Ruby)
    example/example.rb
