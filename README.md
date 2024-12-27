# FedEx Shipping Details Generator

This Python script helps you create bulk FedEx shipping labels from an address sheet in Excel.

## Getting Started

### Prerequisites

Make sure you have Python installed, then run this to install the required libraries:

```bash
pip install pandas openpyxl us
```

### Steps


1. **Update the Config File**:
   Open `config.py` and update these lines:
   - `input_file`: Name of your input Excel file (e.g., `AddressForm.xlsx`).
   - `target_date`: Date to filter the addresses (e.g., `12/16/2024`).

   Example:
   ```python
   input_file = "AddressForm.xlsx"
   output_file = "FedExShippingDetails.xlsx"
   target_date = "12/16/2024"
   ```

2. **Prepare Your Input File**:
   Download the New Joiners Address Form Exxcel Sheet

3. **Run the Script**:
   Run the Python script:

   ```bash
   python3 main.py
   ```

   The script will:
   - Filter addresses by `target_date`.
   - Create shipping details for each address.
   - Save the details in the output file you specified.

5. **Check the Output**:
   The script will create an Excel file with shipping details and open it automatically.

---

## Summary

1. Update `config.py` with your file names and date.
2. Prepare your input Excel file.
3. Run the script to generate shipping details.
