# Risk Profit Analysis - Excel to Database Automation

## Project Overview

This project is designed to automate the processing of Excel data for risk and profit analysis. The system reads structured Excel files, performs calculations, and stores the results in a SQL database. The main purpose is to streamline the import, transformation, and storage of Excel data to support detailed analysis and reporting on risk and profit metrics.

## Key Features

- **Excel File Processing**: Reads data from Excel files and converts it into structured data.
- **Database Integration**: Stores the processed data into a SQL database for further analysis.
- **Automated Updates**: Tracks changes in Excel files and updates the database accordingly.
- **Time-Based Data Tracking**: Handles data updates based on last modification time to ensure only the latest changes are recorded.

## Use Case Scenarios

This template can be used in various scenarios, including but not limited to:

- **Financial Analysis**: Automated data processing for financial forecasts and risk evaluations.
- **Operational Reporting**: Generating periodic reports from operational data stored in Excel files.
- **Data Analytics**: Importing large sets of data into a database for analytics and visualization.
- **Risk Assessment**: Automating the calculation of risk metrics from various sources stored in Excel.

## Getting Started

### Prerequisites

- **.NET Core SDK** installed on your system
- **SQL Server** or another supported database
- **Excel files** formatted as per the predefined template for processing
- **EF Core** (Entity Framework Core) for database management
- **EPPlus** for Excel file handling

### Installation

1. **Clone the Repository**:
   
   ```bash
   git clone https://github.com/Kenterum/Excel_to_Database.git
   cd Excel_to_Database
   ```

2. **Configure Database**:
   
   Update the `appsettings.json` file with your SQL Server connection string:

   ```json
   {
     "ConnectionStrings": {
       "DefaultConnection": "your-sql-connection-string"
     }
   }
   ```

3. **Setup Excel Template**:

   Ensure your Excel files follow the correct format for seamless processing. The folder structure should follow:

   ```
   /RootFolder
     /Month Year
       /Week_1
       /Week_2
       ...
   ```

### Build and Run

1. **Restore Packages**:
   
   ```bash
   dotnet restore
   ```

2. **Build the Project**:
   
   ```bash
   dotnet build
   ```

3. **Run the Application**:
   
   ```bash
   dotnet run
   ```

### Usage

1. **Monitoring Excel Files**:
   
   Place the Excel files in the designated folders. The system will automatically detect updates and process the data accordingly.

2. **Database Updates**:
   
   The system keeps track of changes based on the last modified date of the files and updates the database with the latest data.

### Troubleshooting

- **Database Connection Issues**:
  
  Ensure the connection string in `appsettings.json` is correct and the SQL Server instance is running.
  
- **Excel File Parsing Errors**:
  
  Make sure the Excel files follow the required structure and format.
  
- **Time Zone Differences**:
  
  If there are discrepancies in file modification dates, verify the server and system time zones are correctly configured.

## Contributing

Contributions are welcome! Please follow the standard GitHub process:

1. Fork the repository.
2. Create a new branch.
3. Make your changes.
4. Submit a pull request.

## License

This project is licensed under the MIT License.

## Contact

For any queries or support, please reach out:

- **Email**: suleymammadov@gmail.com
- **GitHub**: [Kenterum](https://github.com/kenterum)

