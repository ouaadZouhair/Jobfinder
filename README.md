# 💼 Morocco Job Finder

A Python web application for searching job opportunities across multiple cities in Morocco using **JobSpy** and **Streamlit**.

Search jobs by keywords, cities, platforms, and publication date. View the results in an interactive dashboard and export them to a formatted Excel file.

## 🚀 Features

- 🔎 Search jobs using multiple keywords.
- 📍 Search across multiple Moroccan cities.
- 🌐 Select multiple job platforms, including Indeed, LinkedIn, Google Jobs, and other JobSpy-supported sources.
- 📅 Filter jobs by publication date.
- ⬇️ Sort job offers from newest to oldest.
- 🗂️ Automatically remove duplicate job offers.
- 📊 View results in an interactive table.
- 📥 Export results to Excel (`.xlsx`).
- 🔗 Access job application links directly.

## 🛠️ Technologies

| Technology | Purpose |
|---|---|
| Python | Main programming language |
| Streamlit | Web application interface |
| JobSpy | Job listing collection |
| Pandas | Data processing and filtering |
| OpenPyXL | Excel generation and formatting |

## 📋 Prerequisites

Before installation, make sure you have:

- Python 3.10 or newer
- pip (Python package manager)
- Git
- An internet connection

Check your Python version:

```bash
python3 --version
```

## 📦 Installation

### 1. Clone the repository

Clone the project using SSH:

```bash
git clone git@github.com:ouaadZouhair/Jobfinder.git
```

Navigate to the project:

```bash
cd Jobfinder
```

Alternatively, clone using HTTPS:

```bash
git clone https://github.com/ouaadZouhair/Jobfinder.git
```

### 2. Create a virtual environment

On Linux or macOS:

```bash
python3 -m venv venv
source venv/bin/activate
```

On Windows:

```powershell
py -m venv venv
.\venv\Scripts\Activate.ps1
```

### 3. Install dependencies

Install all required Python libraries:

```bash
python -m pip install -r requirements.txt
```

If the `requirements.txt` file is not available, install the libraries manually:

```bash
python -m pip install streamlit python-jobspy pandas openpyxl
```

### 4. Start the application

Run:

```bash
streamlit run app.py
```

Open your browser at:

http://localhost:8501

## 📖 How to use

1. Enter one or multiple job keywords, separated by commas or new lines.
2. Select the Moroccan cities you want to search.
3. Choose one or multiple supported job platforms.
4. Select a publication date range.
5. Set the number of requested results per search.
6. Click **Search Jobs**.
7. View and filter the collected job offers.
8. Click **Download Excel** to export your results.

## 📁 Project Structure

```text
Jobfinder/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
└── venv/              # Local virtual environment
```

## 📊 Excel Export

The generated Excel workbook includes:

- Job platform
- Job title
- Company name
- Location
- Publication date
- Job type
- Minimum and maximum salary
- Currency
- Application URL
- Search keyword
- Search city

Excel files include formatted headers, sortable tables, filters, and clickable application links.

## ⚠️ Important Notes

- Job availability depends on the selected platform and location.
- Some platforms restrict automated scraping. Use JobSpy only where the collection method is authorized.
- Publication dates may not include exact posting times.
- Search results may contain incomplete information.
- Avoid excessive automated requests to recruitment websites.

## 👨‍💻 Author

**Zouhair Ouaâd**

GitHub: [ouaadZouhair](https://github.com/ouaadZouhair)

## 📄 License

No license has been assigned yet. See the repository for any future licensing information.