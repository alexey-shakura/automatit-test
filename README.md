# Automatit Test Project

This is a Node.js/Express application that processes Excel files containing invoice data.

## Prerequisites

- Node.js (v14 or higher)
- npm (comes with Node.js)

## Installation

1. Clone the repository
2. Install dependencies:
```bash
npm install
```

## Running the Project

### Development Mode
To run the project in development mode with hot-reload:
```bash
npm run dev
```

The server will start on port 3000 by default. You can change this by setting the `PORT` environment variable.

## API Endpoints

### Upload Excel File
Upload an Excel file containing invoice data.

**Endpoint:** `POST /upload`

**Headers:**
- `Content-Type: multipart/form-data`

**Parameters:**
- `file`: Excel file (.xlsx or .xls)
- `invoicingMonth`: String in YYYY-MM format (e.g., "2024-03")

**Example using curl:**
```bash
curl -X POST \
  http://localhost:3000/upload \
  -H 'Content-Type: multipart/form-data' \
  -F 'file=@/path/to/your/file.xlsx' \
  -F 'invoicingMonth=2024-03'
```

**Response Format:**
```json
{
  "invoicingMonth": "2024-03",
  "currencyRates": {
    "USD": 1.0,
    "EUR": 0.92
  },
  "invoicesData": [
    {
      "Customer": "Example Corp",
      "Cust No": "12345",
      "Project Type": "Development",
      "Quantity": 1,
      "Price Per Item": 1000,
      "Item Price Currency": "USD",
      "Invoice Total Price": 1000,
      "Invoice Currency": "USD",
      "Status": "Ready",
      "Invoice Total": 1000,
      "validationErrors": []
    }
  ]
}
```

## Error Handling

The API will return appropriate error messages for:
- Invalid file format (only Excel files are allowed)
- Missing or invalid invoicing month
- Missing file
- Invalid data in the Excel file
- Currency rate issues 