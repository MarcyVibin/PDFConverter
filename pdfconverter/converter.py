from subprocess import check_output
import win32com.client
import img2pdf
from PyPDF2 import PdfMerger
from pathlib import Path

class Converter:
  class DatatypeConverters:
    def __init__(self):
      pass

    def excel_to_pdf(input: Path, output: Path, sheetnumber=(-1)):
      """Converts an excel file to a pdf file.

      Args:
          input (Path): Path of file.
          output (Path): Output path of file.
          sheetnumber (int, optional): Index of the sheet you want to convert (Index starts with 1). Defaults to all sheets.
      """
      def checkForValueError(sheetnumber: int, amount_sheets: int) -> None:
        #get amount of sheets
        if sheetnumber == 0:
          raise ValueError("Index Error. Sheet index starts with 1.")
        elif not (sheetnumber == -1 or (1 <= sheetnumber <= amount_sheets)):
          raise ValueError("The given sheet index is not valid.")
        return

      def decideOrientation(ws):
        # Check the dimensions of the used range
        used_range = ws.UsedRange
        width = used_range.Width  # Width of used range in points
        height = used_range.Height  # Height of used range in points

        # Determine orientation based on width and height
        if width > height:
          return 2  # Landscape
        else:
          return 1  # Portrait

      def handleSheetParam(ws):
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.CenterHorizontally = True
        ws.PageSetup.FitToPagesTall = False
        ws.PageSetup.Orientation = decideOrientation(ws)
        return

      try:
        excel = win32com.client.DispatchEx("Excel.Application")

        #opening excel file
        excel.Visible = False
        wb = excel.Workbooks.Open(str(input), ReadOnly=True)
        amount_sheets = wb.Sheets.Count
        
        #checks if the sheet exists
        checkForValueError(sheetnumber, amount_sheets)

        #export all sheets
        if sheetnumber == (-1):
          for ws in wb.Sheets:
            handleSheetParam(ws)

          wb.ExportAsFixedFormat(0, str(output))
        #export given sheet
        else:
          ws = wb.Worksheets(sheetnumber)
          handleSheetParam(ws)
          ws.ExportAsFixedFormat(0, str(output))
      except Exception:
        raise
      #always close file
      finally:
        if wb:
            wb.Close(False)
        if excel:
            excel.Quit()
      return
    
    def word_to_pdf(input: Path, output: Path):
      """Converts an word file to a pdf file.

      Args:
          input (Path): Path of file.
          output (Path): Output path of file.
      """
      try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(str(input))
        #pdf file format
        wdFormatPDF = 17
        doc.SaveAs(str(output), FileFormat=wdFormatPDF)
      except Exception:
        raise
      finally:
        if doc:
          doc.Close(False)
        if word:
          word.Quit()
      return

    def ppt_to_pdf(input: Path, output: Path):
      """Converts an powerpoint file to a pdf file.

      Args:
          input (Path): Path of file.
          output (Path): Output path of file.
      """
      try:
        ppt = win32com.client.Dispatch("Powerpoint.Application")
        #ppt.Visible = False

        powerpt = ppt.Presentations.Open(str(input), WithWindow=False)
        #pdf file format
        wdFormatPDF = 32
        powerpt.SaveAs(str(output), FileFormat=wdFormatPDF)
      except Exception:
        raise
      finally:
        if powerpt:
          powerpt.Close()
        if ppt:
          ppt.Quit()
      return

    def img_to_pdf(input: Path, output: Path):
      """Converts an image file to a pdf file.

      Args:
          input (Path): Path of file.
          output (Path): Output path of file.
      """
      with open(str(output), "wb") as f:
        f.write(img2pdf.convert(str(input)))
      return

    #function works but style isn't preserved the way intended
    """def code_to_pdf(input: Path, output: Path, lexer):
      try:
        # Read C++ source code
        with input.open('r', encoding='utf-8') as f:
          code = f.read()

        # Syntax highlight to HTML
        formatter = HtmlFormatter(style='monokai', noclasses=True, linenos='table')
        html_code = highlight(code, lexer, formatter)

        injected_css = '''
        <style>
          body {
            background-color: #1e1e1e !important;
            color: #d4d4d4 !important;
            font-family: "Consolas";
            font-size: 12pt;
            padding: 1em;
          }
          pre {
            white-space: pre-wrap;  /* Allow wrapping */
            word-wrap: break-word;
          }
        </style>
        </head>
        '''

        # Inject CSS into the HTML (before </head>)
        html_code = html_code.replace('</head>', injected_css)
        HTML(string=html_code).write_pdf(str(output))

      except Exception:
        raise """

  def merge_pdfs(directory, output="finshed.pdf") -> None:
    """Merges all pdf files in the same directory.

    Args:
      directory (Path): Directory with the pdf files.
      output (str, optional): Path of the output file. Defaults to "finshed.pdf".
    """

    file_list = []
    for path in Path(directory).iterdir():
      file_list.append(path)

    merger = PdfMerger()
    for path in file_list:
      merger.append(str(path))
    merger.write(output)
    merger.close()
    return
  
  def handle_output_path(output: Path) -> Path:
    """Makes sure file, isn't overwritten but renamed into a copy.

    Args:
      output (Path): Path to be handled. 

    Returns:
      Path: The new path. 
    """
    while (output.is_file()):
      path_suffix = output.suffix
      string_output = str(output)
      try:
        index1 = string_output.index("(") + 1
        index2 = string_output.index(")")
        current_copy = int(string_output[index1:index2])
        string_output = string_output.replace(f"_({current_copy})", f"_({current_copy + 1})")
      except ValueError:
        string_output = (string_output.strip(path_suffix)) + (f"_(1){path_suffix}")

      output = Path(string_output)
    return output

  def checkForExistingPDF(path: Path) -> int:
    """Checks if there already exists a pdf file for the given filename.

    Args:
        input (Path): Path of file.

    Returns:
        int: 1 for exists, 0 for doesnt.
    """
    possible_pdf_path: Path = Path(str(path).strip(path.suffix) + ".pdf")
    if possible_pdf_path.is_file():
      return 1
    else:
      return 0

  def convert(input: Path, output: Path) -> None:
    """Convert a file into a pdf.

    Args:
      input (Path): Path of the file to convert.
      output (Path): Path of the new file.
    """
    
    output = Converter.handle_output_path(output)
    if input.is_file():
      case_type = input.suffix
      #if (Converter.checkForExistingPDF(input) and input.suffix != ".pdf"):
      #  case_type = "exists" #skip to not convert case
      if input.name[0] == "~":
        case_type = "ignore"
      #execute Case
      print("Converting (copying):", input.name)
      match case_type:
        case ".png" | ".jpg" | ".jpeg":
          Converter.DatatypeConverters.img_to_pdf(input, output)
          print("Finished converting:", input.name, "\n")
          return
        case ".xlsx":
          Converter.DatatypeConverters.excel_to_pdf(input, output)
          print("Finished converting:", input.name, "\n")
          return
        case ".pptx":
          Converter.DatatypeConverters.ppt_to_pdf(input, output)
          print("Finished converting:", input.name, "\n")
          return
        case ".docx":
          Converter.DatatypeConverters.word_to_pdf(input, output)
          print("Finished converting:", input.name, "\n")
          return
        #case ".cpp" | ".hpp":
        #  Converter.DatatypeConverters.code_to_pdf(input, output, pygments.lexers.CppLexer())
        #  print("Finished converting:", input.name, "\n")
        #  return
        #case ".c" | ".h":
        #  Converter.DatatypeConverters.code_to_pdf(input, output, pygments.lexers.CLexer())
        #  print("Finished converting:", input.name, "\n")
        #  return
        #case ".py" | ".pyw":
        #  Converter.DatatypeConverters.code_to_pdf(input, output, pygments.lexers.PythonLexer())
        #  print("Finished converting:", input.name, "\n")
        #  return
        case ".pdf":
          check_output(f'copy "{str(input)}" "{str(output)}"', shell=True)
          print("Finished converting:", input.name, "\n")
          return
        case "ignore":
          print("~file will be skipped. (File that is currently opened or being opened by the program)")
          return
        case _:
          print(input.name, ", was not converted.\n")
          return
    return
