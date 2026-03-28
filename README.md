# 📈 AI-Powered Indian Stock Portfolio Manager

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-Deployed-FF4B4B.svg)
![Pandas](https://img.shields.io/badge/Data-Pandas%20%7C%20NumPy-150458.svg)

An end-to-end data engineering pipeline and web application that automates the tracking, analysis, and reporting of Indian stock portfolios (NSE/BSE). Built with Object-Oriented Programming (OOP) principles, this tool extracts live market data, generates algorithmic investment advice, and automatically emails a professionally formatted Excel report to the user.

🌐 **[Live Demo: Try the App Here](https://stocksportfoliomanager.streamlit.app/)**

## ✨ Core Features

* **Live Market Data Integration:** Fetches real-time stock prices using the `yfinance` API.
* **Algorithmic 'AI' Advisor:** Calculates portfolio ROI against the NIFTY 50 benchmark and flags stocks for taking profits (>20% ROI) or triggering stop-losses (<-15% ROI).
* **Automated Email Delivery:** Uses `smtplib` to securely email a daily summary, including a generated Sector Distribution Pie Chart (`matplotlib`).
* **Dynamic Excel Formatting:** Utilizes `openpyxl` to auto-format the exported report (centered alignment, auto-adjusted column widths, currency/percentage formatting, and conditional color scales for ROI).
* **Defensive Programming:** Features robust error handling, RegEx email validation, and automated data cleaning for user-uploaded CSV/XLSX files.

## 🛠️ Tech Stack & Architecture

* **Frontend / Hosting:** Streamlit, Streamlit Cloud Secrets (for secure credential management)
* **Data Processing:** Pandas, NumPy (Vectorized operations for performance over loops)
* **Data Extraction:** `yfinance`
* **Reporting & Visuals:** Matplotlib (Headless Agg backend), `openpyxl`, Python `email.mime`
* **Architecture:** Structured around a highly cohesive, modular `PortfolioManager` class to ensure scalability and maintainability.

## 🚀 How to Run Locally

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/your-username/portfolio_manager.git](https://github.com/your-username/portfolio_manager.git)
   cd portfolio_manager

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt

3. **Set up local secrets:**

   Create a .streamlit/secrets.toml file in the root directory and add your Gmail App Password:
   ```Ini, TOML
   GMAIL_USER = "your_email@gmail.com"
   GMAIL_PASSWORD = "your_16_digit_app_password"

5. **Run the app:**
   ```bash
   streamlit run app.py

## 📊 Expected Input Format

The application accepts .xlsx or .csv files. The data must contain the following exact column headers (order does not matter):
* `Ticker` (e.g., RELIANCE.NS, TCS.NS)
* `Quantity` (integer)
* `Buy_Price` (float)
* `Sector` (string)

## 👨‍💻 Author

**Sujal Shrivastava**
| B.Tech (Computer Science & Business Systems)
* [LinkedIn](https://www.linkedin.com/in/sujal-shrivastava-355470250/)
