# TikTok-Product-Scraper-with-Selenium-and-Excel-Input
Automate TikTok Shop scraping with Python & Selenium. Extract product title, price, and sales from URLs in Excel, then export to Excel. Perfect for product research, competitor monitoring, and e-commerce data automation.

This Python script automates the process of scraping product details from TikTok Shop. It reads product URLs from an Excel file and extracts the following information:

- Product Title
- Price
- Units Sold
- Scraping Date

The data is saved into a new Excel file.

---

## 📦 Features

- ✅ Reads product URLs from Excel file (supports multiple sheets)
- ✅ Collects title, price, and sales data
- ✅ Stores results in a new Excel file
- ✅ Uses Selenium for full browser automation

---

## 🚀 How to Use

### 1. **Install Dependencies**

Make sure you have Python installed, then install the required packages:

pip install pandas selenium openpyxl

### 2. **Prepare the Excel File**
Create an Excel file (.xlsx) with the following structure:

URL	ARTICLE
https://...tiktok-product...	SKU1234
https://...tiktok-product...	SKU5678

The columns URL and ARTICLE must be present in each sheet.

### 3. **Run the Script**

python tiktok_scraper.py

A file dialog will appear—select your Excel file. The script will then process each URL and extract the relevant data.


📝 Output
The script will generate a file named like:

tiktok_24 Jun 25.xlsx
Containing the following columns:

- URL
- Article
- Title
- Price
- Total Sold
- Current Date

❗ Notes
TikTok's HTML structure may change frequently. If scraping fails, class names may need updating in the script.

Be responsible with your scraping speed. Avoid running too many requests in a short time to prevent getting blocked.

🧑‍💻 Author
Developed by dewarakaks — feel free to fork, improve, or adapt the code for your own needs!

Let me know if you want me to:
- Include a sample Excel file for users
- Package this as an installable tool
- Convert it into a headless or multi-threaded scraper

I'm happy to help polish this further!
