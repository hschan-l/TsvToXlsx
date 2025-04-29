#include <iostream>
#include <fstream>
#include <sstream>
#include <string>
#include <map>
#include <algorithm>
#include "miniz.h" 

// Function to trim whitespace from a string
std::string trim(const std::string& str) {
    size_t first = str.find_first_not_of(" \t\n\r");
    if (first == std::string::npos) return "";
    size_t last = str.find_last_not_of(" \t\n\r");
    return str.substr(first, last - first + 1);
}

// Function to escape XML special characters
std::string escapeXml(const std::string& value) {
    std::string result = value;
    size_t pos = 0;
    while ((pos = result.find('&', pos)) != std::string::npos) {
        result.replace(pos, 1, "&amp;");
        pos += 5;
    }
    pos = 0;
    while ((pos = result.find('<', pos)) != std::string::npos) {
        result.replace(pos, 1, "&lt;");
        pos += 4;
    }
    pos = 0;
    while ((pos = result.find('>', pos)) != std::string::npos) {
        result.replace(pos, 1, "&gt;");
        pos += 4;
    }
    pos = 0;
    while ((pos = result.find('"', pos)) != std::string::npos) {
        result.replace(pos, 1, "&quot;");
        pos += 6;
    }
    pos = 0;
    while ((pos = result.find('\'', pos)) != std::string::npos) {
        result.replace(pos, 1, "&apos;");
        pos += 6;
    }
    return result;
}

// Function to convert column index to Excel column reference (A, B, ..., Z, AA, AB, ...)
std::string getColumnReference(int col) {
    std::string colRef;
    do {
        colRef.insert(0, 1, 'A' + (col % 26));
        col = col / 26 - 1;
    } while (col >= 0);
    return colRef;
}

// Function to read the input file (assumes UTF-8)
std::string readFile(const std::string& filename) {
    std::ifstream file(filename, std::ios::binary);
    if (!file.is_open()) {
        return "";
    }
    std::ostringstream content;
    content << file.rdbuf();
    file.close();
    std::cout << "Read file as UTF-8.\n";
    return content.str();
}

// Function to create sheet1.xml from tab-separated content
std::string createSheetXml(const std::string& content) {
    std::ostringstream sheetData;
    sheetData << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)""\n";
    sheetData << R"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)""\n";
    sheetData << "  <sheetData>\n";

    std::istringstream reader(content);
    std::string line;
    int rowNum = 1;
    while (std::getline(reader, line)) {
        std::string trimmedLine = trim(line);
        if (trimmedLine.empty()) continue; // Skip empty lines
        sheetData << "    <row r=\"" << rowNum << "\">\n";
        std::istringstream lineStream(trimmedLine);
        std::string cell;
        int col = 0;
        while (std::getline(lineStream, cell, '\t')) {
            std::string colRef = getColumnReference(col);
            std::string cellRef = colRef + std::to_string(rowNum);
            std::string escapedValue = escapeXml(trim(cell));
            sheetData << "      <c r=\"" << cellRef << "\" t=\"inlineStr\">\n";
            sheetData << "        <is><t>" << escapedValue << "</t></is>\n";
            sheetData << "      </c>\n";
            col++;
        }
        sheetData << "    </row>\n";
        rowNum++;
    }

    sheetData << "  </sheetData>\n";
    sheetData << "</worksheet>\n";
    return sheetData.str();
}

// Function to create the ZIP file (xlsx) with all required XML files
void createZip(const std::string& zipFilename, const std::string& content) {
    std::map<std::string, std::string> files;

    // 1. [Content_Types].xml
    files["[Content_Types].xml"] = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>)";

    // 2. _rels/.rels
    files["_rels/.rels"] = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";

    // 3. xl/workbook.xml
    files["xl/workbook.xml"] = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>)";

    // 4. xl/_rels/workbook.xml.rels
    files["xl/_rels/workbook.xml.rels"] = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>)";

    // 5. xl/worksheets/sheet1.xml
    files["xl/worksheets/sheet1.xml"] = createSheetXml(content);

    // Create ZIP archive using miniz
    mz_zip_archive zip_archive;
    memset(&zip_archive, 0, sizeof(zip_archive));
    if (!mz_zip_writer_init_file(&zip_archive, zipFilename.c_str(), 0)) {
        std::cerr << "Error: Failed to initialize ZIP file '" << zipFilename << "'.\n";
        return;
    }

    for (const auto& [filename, data] : files) {
        if (!mz_zip_writer_add_mem(&zip_archive, filename.c_str(), data.c_str(), data.size(), MZ_DEFAULT_COMPRESSION)) {
            std::cerr << "Error: Failed to add '" << filename << "' to ZIP.\n";
            mz_zip_writer_end(&zip_archive);
            return;
        }
    }

    if (!mz_zip_writer_finalize_archive(&zip_archive)) {
        std::cerr << "Error: Failed to finalize ZIP archive.\n";
    }
    mz_zip_writer_end(&zip_archive);
}

int main(int argc, char* argv[]) {
    if (argc < 2) {
        std::cout << "Usage: " << argv[0] << " <input_file.txt>\n";
        return 1;
    }

    std::string inputFile = argv[1];
    std::ifstream fileCheck(inputFile);
    if (!fileCheck.good()) {
        std::cout << "Error: File '" << inputFile << "' not found.\n";
        return 1;
    }
    fileCheck.close();

    std::string outputXlsx = inputFile.substr(0, inputFile.find_last_of('.')) + ".xlsx";
    std::cout << "Creating Excel file from '" << inputFile << "' with tab separators...\n";

    std::string content = readFile(inputFile);
    if (content.empty()) {
        std::cerr << "Error: Could not read file '" << inputFile << "'. Exiting.\n";
        return 1;
    }

    try {
        createZip(outputXlsx, content);
    }
    catch (const std::exception& e) {
        std::cerr << "Error processing file: " << e.what() << "\n";
        return 1;
    }

    std::cout << "Excel file saved as: " << outputXlsx << "\n";
    std::cout << "Done. Open " << outputXlsx << " in Excel to confirm.\n";
    return 0;
}




