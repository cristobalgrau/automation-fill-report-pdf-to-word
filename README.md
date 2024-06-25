
![image](https://github.com/cristobalgrau/automation-fill-report-pdf-to-word/assets/119089907/c9db6e26-d217-44c6-ae77-948873aaa9b0)


# Automation: Fill out a Table in a Word Document from a PDF information


In my current job, we typically receive a Job Order containing all the data about the vessel we are going to attend. This Job Order is crucial for understanding the issues at hand and preparing for troubleshooting and repairs. The Job Order comes in a PDF file, and we need to extract specific information from it to populate our Service Report, a Word document, with the necessary details in predefined table cells.

To streamline this process, I developed a Python script that automates the extraction of data from the PDF and populates the corresponding fields in the Word document. This automation saves significant time and reduces the monotonous task of manually transferring data, allowing us to focus more on the actual troubleshooting and solution development.

## Library Used:

- **PyPDF2**: This library is used for reading and extracting text from PDF files. It simplifies the process of accessing and parsing PDF content.
	[PyPDF2 Documentation](https://pypdf.readthedocs.io/en/stable/)
- **python-docx**: This library is used for creating and updating Microsoft Word (.docx) files. It allows us to manipulate document content programmatically, including text replacement and formatting.
	[python-docs Documentation](https://python-docx.readthedocs.io/en/latest/index.html)

## Key Functions:

### Working with the PDF document

The information needed comes in the following part of the Order (Information was erased for privacy):

![image](https://github.com/cristobalgrau/automation-fill-report-pdf-to-word/assets/119089907/24066aec-9d27-4150-b12c-42b757b700f2)


I used the following functions to extract and clean the information:


- `extract_text_from_pdf(pdf_path)`
```python
def extract_text_from_pdf(pdf_path):
	with open(pdf_path, "rb") as pdf:
		reader = pypdf.PdfReader(pdf)
		page = reader.pages[0].extract_text()
	return page.splitlines()
```

This function opens and reads the PDF file (`pdf_path`), extracts text from the first page of the PDF, and splits the extracted text into lines converting them into a list. Then return this list


- `clean_text_list(text_list)`
```python
def  clean_text_list(text_list):
	text_list = [line for line in text_list if line.strip()] # Remove empty lines
	return  text_list[4:12] # Get relevant lines
```

This function cleans the list by removing the empty lines, and extracts the relevant lines, in this case lines from 4 to 11, which contain the information needed.


- `extract_values(new_list)`
```python
def extract_values(new_list):
    # Extract and clean values
    attendance_date = new_list[0].replace("<XXXXXXXXXXXXXXXX>  Attendance Date", "").strip()
    location = new_list[1].replace("Attendance Location", "").strip()
    ticket = new_list[2].replace("Customer", "").strip()
    antenna_serial = new_list[6].replace("Antenna Serial", "").strip()
    antenna_model = new_list[7].replace("Antenna type", "").strip()

    # Obtain WO and Case
    temp = new_list[3].replace("Click or tap here to enter text.", "").replace("Case / WO", "").lstrip()
    wo_complete = temp.split()[0]
    case = temp.replace(wo_complete, "").strip()
    case = case.strip(" ()")
    wo = wo_complete.replace("WO-", "000")

    # Obtain Vessel Name
    temp = new_list[4].replace("Request Date", "")
    request_date = temp.split()[0]
    vessel = temp.replace(request_date, "").lstrip().replace("Vessel Name", "").strip()

    placeholders = {
        "attendance_date": attendance_date,
        "[location]": location,
        "[ticket]": ticket,
        "[wo]": wo,
        "[case]": case,
        "[vessel]": vessel,
        "[antenna_serial]": antenna_serial,
        "[antenna_model]": antenna_model,
    }

    file_name = f"{wo_complete}_Field Service Report_{vessel}_date.docx"

    return placeholders, file_name
```

This function extracts and cleans specific values from the list, constructs a dictionary of placeholders (matching the placeholders in the Word document), and saves the respective values for those placeholders. Finally, it generates a name for the new Word document created with the populated information.


### Working with the WORD document

The part of the Word document that needs to be populated is the following (Some information was covered for privacy)

![image](https://github.com/cristobalgrau/automation-fill-report-pdf-to-word/assets/119089907/48b8c7c8-22cb-4be0-8b31-5e59a6f755ea)

The following functions were created to edit the word and populate it with the information extracted from previous steps:

- `replace_text_and_align(cell, placeholder, replacement)`
```python
def replace_text_and_align(cell, placeholder, replacement):
    if cell.text == placeholder:
        cell.text = cell.text.replace(placeholder, replacement)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        print(f"Found {placeholder} and modified!")
```

This function replaces placeholder text in a Word table cell with the specified replacement text and then centers the text within the cell. The placeholders were created in the function `extract_values()`


- `replace_paragraph_text(doc, placeholder, replacement)`
```python
def replace_paragraph_text(doc, placeholder, replacement):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, replacement)
            print(f"Found {placeholder} and modified!")
```

In this function, we replace the placeholder for the `attendance_date` because this value is not inside a table and needs to be treated as a paragraph, following the documentation from `python-docx`


## Main Workflow:

1. **Extract Data from PDF**:
    
    -   The script starts by extracting text from the specified PDF file (`pdf_path`).
    -   The relevant lines are cleaned and specific values are extracted into a dictionary of placeholders.

2. **Populate Word Document**:
    
    -   The script loads the template Word document (`doc_path`).
    -   Sets the default font and size for the document.
    -   Replaces the placeholder for the attendance date in paragraphs.
    -   Iterates through the tables in the document and replaces placeholders with the extracted values, aligning the text as necessary.
    -   Saves the modified document with a dynamically generated file name.

## Result:

![image](https://github.com/cristobalgrau/automation-fill-report-pdf-to-word/assets/119089907/b9d3c3e8-14aa-44c6-867b-f36bc37f1663)


The final Doc can't be shown due to sensitive information


## Benefits of this Automation

1.  **Efficiency**: It significantly reduces the time spent on manual data entry, allowing staff to focus on more critical tasks.
2.  **Accuracy**: By automating data extraction and insertion, the chances of human error are minimized, ensuring that the reports are accurate and consistent.
3.  **Consistency**: The script ensures that all reports are formatted uniformly, which is important for maintaining professional standards and ease of review.
4.  **Scalability**: As the volume of Job Orders increases, the automation can handle more documents without requiring additional manpower.

By implementing this Python automation, we enhance our workflow, reduce tedious tasks, and improve overall productivity and accuracy in our documentation process.

