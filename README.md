# PDF_Word_Searcher
A python based project to search for a specific word and return a list containing the name of PDF files where the word we search for is contained inside it, each filename in the list is hyperlinked with the exact path of the file itself. It can process documents, such as word(docx,doc), presentation(pptx), excel(xlsx), log, textfile (txt).
Module used:<br /> 
- "tkinter"(GUI)<br />
- "pymupdf(fitz)" for PDF file processing<br />
- "os" (directory and saving file)<br />
- "nuitka" to turn the python script into executable application with EXE format<br />
- Pandas
- python-pptx, to process presentation document<br />
- python-docx, to process word document<br />
- win32com.client, to convert .doc into .docx<br />
- openpyxl, to process excel document<br />
- report-lab, to create a list of filenames into a PDF file<br />
- urlib.parse, to import "quote"
- Need NotoSansTC font pack downloaded and put it into same file as your program (Provided below)


![Screenshot (949)](https://github.com/user-attachments/assets/4cf3a55e-9b96-4403-bf49-d91748e8f4b0)

I made this application to fulfill my senior request ( im doing a part time job in my univ library) to make a PDF search engine that wont connect to the internet for data protection.

For the nuitka, i use the following code to convert my script (im using visual studio code), run it from the terminal:
nuitka --standalone --onefile --windows-disable-console --enable-plugin=tk-inter --output-filename=PDF_Searcher.exe meili.py



