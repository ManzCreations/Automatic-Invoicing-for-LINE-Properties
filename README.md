# Airbnb and VRBO Invoice Report Generator

This Python script automates the generation of comprehensive monthly invoice reports for property managers and owners who list their properties on Airbnb and VRBO. It simplifies financial record-keeping by consolidating rental transactions into detailed invoices, journal entries, checks, and sales tax calculations. Additionally, it maintains a dynamic client database that updates automatically with each run.

## Features

- **Invoice Generation**: Automatically creates detailed invoices for each client based on Airbnb and VRBO transactions.
- **Journal Entries and Checks**: Generates journal entries and checks for accurate financial tracking.
- **Sales Tax Calculations**: Computes sales tax obligations based on rental income.
- **Client Database Management**: Keeps a real-time updated database of clients and their transaction histories.
- **Monthly Reports**: Designed to run monthly, supporting property managers in keeping consistent financial records.

## Getting Started

### Prerequisites

Before running this script, ensure you have Python 3.x installed on your system. Additionally, the following Python packages are required:

- pandas
- numpy
- openpyxl
- xlsxwriter
- datetime
- re (regular expressions)

You can install these packages using pip:
```
pip install pandas numpy openpyxl xlsxwriter
```

## Installation
Clone the repository to your local machine:
```
git clone https://github.com/ManzCreations/Automatic-Invoicing-for-LINE-Properties.git
```
Navigate to the cloned repository:
```
cd Automatic-Invoicing-for-LINE-Properties
```
Install the required Python packages:
```
pip install -r requirements.txt
```
## Usage
To run the script, execute the following command in the terminal:
```
python invoice_generator.py
```
Make sure to adjust the script's file paths and configurations based on your specific setup and requirements.

## Contributing
Contributions to improve the script are welcome. Please follow these steps to contribute:

1. Fork the repository.
2. Create a new branch (git checkout -b feature/AmazingFeature).
3. Commit your changes (git commit -m 'Add some AmazingFeature').
4. Push to the branch (git push origin feature/AmazingFeature).
5. Open a pull request.

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments
Thanks to Airbnb and VRBO for providing the transaction data formats.
This script is intended for educational purposes and practical applications by property managers.
