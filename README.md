<h1>PDF to Excel Automation Script</h1>

<p>This Python script automates the process of extracting data from a PDF file and writing it into a pre-defined Excel template. The primary modules used include <code>pdfplumber</code> for PDF extraction, <code>pandas</code> for data manipulation, and <code>xlwings</code> for interacting with the Excel file.</p>

<h2>Features</h2>
<ul>
    <li>Extracts data from a PDF file with specific table structures.</li>
    <li>Appends the extracted data into an Excel template at specified locations.</li>
    <li>Formats the data into proper rows and columns.</li>
    <li>Automatically fills in dates, project numbers, and calculates the total order value.</li>
    <li>Converts numeric totals into words and writes them to the Excel sheet.</li>
    <li>Performs cleanup tasks such as removing unused rows and columns in the Excel template.</li>
</ul>

<h2>Installation</h2>
<p>To run this script, you'll need the following Python libraries installed:</p>

<pre>
pip install pdfplumber pandas xlwings inflect openpyxl
</pre>

<h2>How to Use</h2>

<ol>
    <li>Place the PDF file you want to process in the same directory as this script, and name it <code>1.pdf</code>.</li>
    <li>Ensure the Excel template file (<code>order_template.xls</code>) is in the same directory.</li>
    <li>Run the script. The script will extract table data from the PDF and populate it into the specified cells of the Excel template.</li>
    <li>The filled Excel file will be saved as <code>B2409000313.xls</code>.</li>
</ol>

<h2>Code Overview</h2>

<p>The code follows these steps:</p>
<ol>
    <li>It reads a PDF file named <code>1.pdf</code> using the <code>pdfplumber</code> module.</li>
    <li>Extracts relevant columns ("Supplier Code", "Qty", and "Unit Cost") from each page of the PDF and appends them into a pandas DataFrame.</li>
    <li>Connects to an Excel template using the <code>xlwings</code> module and writes data into specified rows and columns.</li>
    <li>Fills out additional details such as the date, project number, and converts the total value into words using the <code>inflect</code> module.</li>
    <li>Performs post-processing on the Excel file by deleting unnecessary rows and columns.</li>
</ol>

<h2>File Structure</h2>
<p>The expected file structure for this script is as follows:</p>
<div class="highlight">
<pre>
project-folder/
│
├── 1.pdf                   # PDF file to be processed
├── order_template.xls       # Excel template
└── script.py                # This Python script
</pre>
</div>

<h2>Customization</h2>

<p>Some sections of the code may need customization based on your requirements:</p>
<ul>
    <li>File paths: If your files are not in the current working directory, update the paths to match your folder structure.</li>
    <li>Excel template: The script assumes a specific template structure. Modify the cell references in the code to match your Excel file's layout.</li>
</ul>

<h2>Notes</h2>

<ul>
    <li>Ensure that the PDF file contains a table with the columns "Supplier Code", "Qty", and "Unit Cost" for correct extraction.</li>
    <li>The script uses the <code>inflect</code> library to convert numeric totals into words.</li>
</ul>

<h2>License</h2>
<p>This project is open-source and available for modification.</p>

</body>
</html>
