# Excel到Word的批量数据填充（基于csv文件的操作）

本项目旨在帮助用户将Excel表格中的数据批量填充到Word文档中，以提高数据处理的效率和准确性。该项目基于Python编程语言，利用CSV文件作为数据源，并借助第三方库`docx`来处理Word文档。

## 项目背景

在日常工作中，我们经常遇到需要将Excel中的数据填充到Word文档中的情况。传统的手动复制粘贴方式耗时且容易出错，尤其是在处理大量数据时。因此，开发一个自动化的批量数据填充工具具有重要意义。

## 技术实现

本项目采用Python编程语言作为开发工具，并依托以下技术实现主要功能：

- 使用`csv`模块读取和解析CSV文件中的数据。
- 借助`docx`库读取和操作Word文档，包括段落、表格等元素的复制、替换和保存。
- 使用`os`模块进行文件夹的创建和管理，确保生成的Word文档被正确保存。
- 利用`copy`模块中的`deepcopy`函数实现模板文档的复制，确保每个生成的Word文档都基于同一个模板。

## 使用方法

1. 准备一个Excel表格，将需要填充到Word文档的数据整理为一列或多列，并将表格另存为CSV文件格式。
2. 准备一个Word文档作为模板，将需要填充数据的位置用占位符标记，占位符可以是任意文本（即CVS文件第一行：
例子：
姓名，身份证号，地址
那么相对应的占位符即为：{{姓名}}与{{身份证号}}与{{地址}}
）
3. 修改Python代码中的文件路径和占位符规则，确保与自己的文件相匹配。
4. 运行Python代码，它将自动读取CSV文件中的数据，根据占位符规则替换模板中的占位符（即CVS文件第一行），并生成填充好数据的Word文档。
5. 默认在代码运行目录建立一个名为（Excel-Cvs-Python-Docx）的文件夹存放Word文档。

## 总结

本项目提供了一个简单而高效的解决方案，帮助用户实现Excel到Word的批量数据填充操作。通过自动化处理，减少了繁琐的手动操作，提高了工作效率，并减少了人为错误的风险。用户可以根据自己的需求和具体情况，灵活调整代码，实现个性化的数据填充过程。

请在GitHub上查看完整项目代码和详细说明。

链接：[https://github.com/guiyi-xuankong/Excel-Cvs-Python-Docx)


# Excel to Word batch data filling (csv file based operation)

This project aims to help users to batch fill data from Excel tables to Word documents to improve the efficiency and accuracy of data processing. The project is based on Python programming language, using CSV files as data source and processing Word documents with the help of a third-party library `docx`.

## Project Background

In our daily work, we often encounter the situation that we need to fill data in Excel into Word documents. The traditional manual copy-and-paste method is time-consuming and error-prone, especially when dealing with large amounts of data. Therefore, it is important to develop an automated batch data filling tool.

## Technical Implementation

This project uses the Python programming language as the development tool and relies on the following technologies to achieve the main functions:

- Reading and parsing data in CSV files with the `csv` module.
- Read and manipulate Word documents with the help of `docx` library, including copying, replacing and saving of paragraphs, tables and other elements.
- Use the `os` module for folder creation and management to ensure that the generated Word documents are saved correctly.
- Use the `deepcopy` function in the `copy` module for template document copying to ensure that each generated Word document is based on the same template.

## How to use

1. Prepare an Excel table, organize the data to be filled into Word documents into one or more columns, and save the table as a CSV file format.
2. Prepare a Word document as a template and mark the position of the data to be filled with a placeholder, which can be any text (i.e. the first line of the CVS file:
Example:
Name, ID number, address
then the corresponding placeholder that is: {{name}} and {{ID number}} and {{address}}
)
3. Modify the file path and placeholder rules in the Python code to make sure they match your own files.
4. Run the Python code, it will automatically read the data in the CSV file, replace the placeholders in the template (i.e. the first line of the CVS file) according to the placeholder rules, and generate a Word document with the data filled.
5. default in the code run directory to create a folder named (Excel-Cvs-Python-Docx) to store Word documents.

## Summary

This project provides a simple and efficient solution to help users achieve Excel to Word batch data filling operations. By automating the process, it reduces tedious manual operations, improves work efficiency, and reduces the risk of human error. Users can flexibly adjust the code according to their needs and specific situations to achieve a personalized data filling process.

Please check out the full project code and detailed description on GitHub.

Link: [https://github.com/guiyi-xuankong/Excel-Cvs-Python-Docx)
