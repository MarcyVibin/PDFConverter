from setuptools import setup, find_packages

setup(
  name="pdfconverter",
  version="0.1.0",
  author="marcyvibin",
  description="A simple Python library for converting documents to PDF",
  packages=find_packages(),
  install_requires=[
    "pywin32",
    "PyPDF2",
    "img2pdf",
  ],
  python_requires=">=3.7",
)