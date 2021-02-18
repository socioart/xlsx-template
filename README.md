# XlsxTemplate

    mvn package assembly:single
    java -jar target/xlsx-template-jar-with-dependencies.jar template.xlsx

    mvn compile exec:java -Dexec.mainClass=com.socioart.XlsxTemplate -Dexec.args='template.xlsx'
