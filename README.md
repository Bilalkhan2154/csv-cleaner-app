✨ DataPolish

DataPolish is a free, browser-based tool that cleans messy CSV and Excel files in one click — no installation, no coding, no technical knowledge required.

Upload a file, choose what to fix, click Clean, and download the polished result along with a plain-English report of everything that was changed.

🔗 Live app: your-streamlit-link-here


🚀 Features


📄 CSV & Excel support — upload .csv or .xlsx, download in the same format
🗑️ Remove duplicate rows
🧩 Remove rows with missing values
📅 Fix inconsistent date formats — auto-detects date columns and standardizes them
🔤 Fix text casing — convert to UPPERCASE, lowercase, or Title Case
📊 Before vs After comparison — see exactly what changed, side by side
📋 Downloadable cleaning report — a plain-English summary of every fix applied
🎨 Clean, professional UI — designed to be usable by anyone, not just developers



🖥️ How to Use (Online)


Open the live app link
Upload your CSV or Excel file
Choose your cleaning options in the sidebar
Click Clean My File Now
Download your cleaned file and cleaning report


That's it — no signup, no installation.


💻 Run Locally

Prerequisites


Python 3.8+


Steps

bash# 1. Clone the repository
git clone https://github.com/yourusername/csv-cleaner.git
cd csv-cleaner

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
streamlit run app.py

The app will open automatically in your browser at http://localhost:8501.


📦 Tech Stack


Streamlit — web app framework
Pandas — data cleaning and manipulation
OpenPyXL — Excel file support



📁 Project Structure

csv-cleaner/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md           # This file


🛠️ Deployment

This app is deployed for free on Streamlit Community Cloud. To deploy your own copy:


Fork or upload this repo to your own GitHub account
Go to share.streamlit.io and sign in with GitHub
Click New app → select your repo → main file app.py → Deploy
Share your generated link with anyone



🤝 Contributing

Suggestions and improvements are welcome. Feel free to open an issue or submit a pull request.


📄 License

This project is open source and available under the MIT License.
