<a id="readme-top"></a>

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![License][license-shield]][license-url]

<h3 align="center">PDFConverter</h3>

<p align="center">
    Convert Excel, Word, PowerPoint and image files into PDF, and merge PDFs, with a few lines of Python.
    <br />
    <br />
    <a href="https://github.com/MarcyVibin/PDFConverter/issues">Report Bug</a>
    &middot;
    <a href="https://github.com/MarcyVibin/PDFConverter/issues">Request Feature</a>
</p>

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#requirements">Requirements</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
  </ol>
</details>

---

## About The Project

**PDFConverter** is a Python library that transforms various file types into **PDF**. It uses Microsoft Office COM automation (via `pywin32`) and a few small libraries to convert:

- Excel (`.xlsx`)
- Word (`.docx`)
- PowerPoint (`.pptx`)
- Images (`.png`, `.jpg`, `.jpeg`)
- Existing PDFs (copied to the new location)

It can also **merge all PDFs in a folder** into a single document, and it never overwrites an existing file: if the output name is taken, the new file gets a copy suffix.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### Built With

* [![Python][Python.org]][Python-url]
* [![Microsoft Office][Office.com]][Office-url]
* [![PyWin32][PyWin32-badge]][PyWin32-url]
* [![PyPDF2][PyPDF2-badge]][PyPDF2-url]
* [![img2pdf][img2pdf-badge]][img2pdf-url]

<p align="right">(<a href="#readme-top">back to top</a>)</p>

---

## Getting Started

### Requirements

- Windows
- Microsoft Office (Excel, Word and PowerPoint installed)
- Python 3.10+
- `pywin32`, `img2pdf` and `PyPDF2`. They are installed automatically with the package.

### Installation

Install with pip:

```sh
pip install git+https://github.com/MarcyVibin/PDFConverter.git
```

Or clone the repository and install it from the folder:

```sh
git clone https://github.com/MarcyVibin/PDFConverter.git
cd PDFConverter
pip install .
```

---

## Usage

```py
from pathlib import Path
from pdfconverter import Converter

# Convert one file; the type is picked from the file extension
Converter.convert(Path("report.xlsx"), Path("report.pdf"))

# Merge every PDF in a folder into one file
Converter.merge_pdfs("pdfs", "merged.pdf")

# Only needed to export a single Excel sheet (the index starts at 1)
Converter.DatatypeConverters.excel_to_pdf(Path("report.xlsx"), Path("sheet2.pdf"), sheetnumber=2)
```

<p align="right">(<a href="#readme-top">back to top</a>)</p>

---

## License

Distributed under the GNU General Public License v3.0 (GPL-3.0). See `LICENSE.txt` for more information.

---

## Contact

Marc Steger - marc.steger06@gmail.com

Project Link: https://github.com/MarcyVibin/PDFConverter

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- Badge Links -->

[Python.org]: https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white
[Python-url]: https://www.python.org/

[Office.com]: https://img.shields.io/badge/Microsoft%20Office-D83B01?style=for-the-badge&logo=microsoft-office&logoColor=white
[Office-url]: https://www.microsoft.com/microsoft-365

[PyWin32-badge]: https://img.shields.io/badge/PyWin32-003366?style=for-the-badge
[PyWin32-url]: https://pypi.org/project/pywin32/

[PyPDF2-badge]: https://img.shields.io/badge/PyPDF2-004C97?style=for-the-badge
[PyPDF2-url]: https://pypi.org/project/PyPDF2/

[img2pdf-badge]: https://img.shields.io/badge/img2pdf-222222?style=for-the-badge
[img2pdf-url]: https://pypi.org/project/img2pdf/

[contributors-shield]: https://img.shields.io/github/contributors/MarcyVibin/PDFConverter.svg?style=for-the-badge
[contributors-url]: https://github.com/MarcyVibin/PDFConverter/graphs/contributors

[forks-shield]: https://img.shields.io/github/forks/MarcyVibin/PDFConverter.svg?style=for-the-badge
[forks-url]: https://github.com/MarcyVibin/PDFConverter/network/members

[stars-shield]: https://img.shields.io/github/stars/MarcyVibin/PDFConverter.svg?style=for-the-badge
[stars-url]: https://github.com/MarcyVibin/PDFConverter/stargazers

[issues-shield]: https://img.shields.io/github/issues/MarcyVibin/PDFConverter.svg?style=for-the-badge
[issues-url]: https://github.com/MarcyVibin/PDFConverter/issues

[license-shield]: https://img.shields.io/github/license/MarcyVibin/PDFConverter.svg?style=for-the-badge
[license-url]: https://github.com/MarcyVibin/PDFConverter/blob/main/LICENSE.txt
