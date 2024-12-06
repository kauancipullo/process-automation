# Process Automation

> **Automates daily OnePage report generation and emailing for stores, including performance indicators and rankings.

## 🚀 Technologies and Libraries Used

- **Python**: Python 3.12.1.
- **Jupyter Notebook**: To view and run notebooks.
- **Pandas**: To handle data easily.
- **PathLib**: To handle files on the computer.
- **win32com.client**: To use Outlook as e-mail sender.

## 📋 Features:
- **Automated OnePage Generation**: Calculates and generates daily and annual performance reports (OnePages) for each store.
- **Email Automation**: Sends OnePage reports via email to store managers with sales data attached.
- **Store-Specific Reports**: Ensures each store manager only receives their own store's data, keeping other store information confidential.
- **Ranking Reports**: Generates daily and annual rankings of stores based on revenue and sends them to management.
- **Management Summary**: Sends a summary email to management highlighting the best and worst-performing stores of the day and year.
- **Backup History**: Saves generated reports in store-specific folders with timestamps to create a backup history of data.
- **HTML Formatting for Symbols**: Uses HTML formatting to display green and red indicators for goal achievement.
- **Data Handling**: Utilizes multiple data files (`Emails.xlsx`, `Vendas.xlsx`, `Lojas.csv`) to handle and process store information.

## 📂 Project Structure
```plaintext
├── process-automation/   # Main directory containing notebooks
    ├── main.ipynb        # Main Notebook
    ├── description.ipynb # Notebook with complete project description
    ├── data/             # Data Directory
    ├── requirements.txt  # Project dependencies
    └── README.md         # This File
```

## ✨ Author
- Name: Kauan Gouveia Cipullo.
- LinkedIn: https://www.linkedin.com/in/kauan-cipullo/
- Email: kaauan_@hotmail.com
