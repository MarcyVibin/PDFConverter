<a id="readme-top"></a>

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![project_license][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]

<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/github_username/file-to-pdf-converter">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a>

<h3 align="center">File to PDF Converter</h3>

  <p align="center">
    Convert Excel, Word, PowerPoint, Images, and PDFs into PDF format with ease.
    <br />
    <a href="https://github.com/github_username/file-to-pdf-converter"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    <a href="https://github.com/github_username/file-to-pdf-converter/issues">Report Bug</a>
    &middot;
    <a href="https://github.com/github_username/file-to-pdf-converter/issues">Request Feature</a>
  </p>
</div>

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
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>

<!-- ABOUT THE PROJECT -->
## About The Project

[![Product Name Screen Shot][product-screenshot]](https://example.com)

This project provides a **Python-based file converter** that can transform various file types into **PDF format**.  
It uses Microsoft Office COM automation (via `pywin32`) and additional libraries to support conversions from:

- Excel (`.xlsx`)
- Word (`.docx`)
- PowerPoint (`.pptx`)
- Images (`.png`, `.jpg`, `.jpeg`)
- Existing PDFs (copied/renamed)

It also supports **merging multiple PDFs** into a single document and handles naming conflicts automatically.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### Built With

* [![Python][Python.org]][Python-url]
* [![Microsoft Office][Office.com]][Office-url]
* [![PyWin32][PyWin32-badge]][PyWin32-url]
* [![PyPDF2][PyPDF2-badge]][PyPDF2-url]
* [![img2pdf][img2pdf-badge]][img2pdf-url]

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- GETTING STARTED -->
## Getting Started

Follow these steps to set up and run the project locally.

### Prerequisites

- Windows OS
- Microsoft Office (Excel, Word, PowerPoint installed)
- Python 3.9+
- Install dependencies:
  ```sh
  pip install pywin32 img2pdf PyPDF2
