import pandas as pd
import fuzzywuzzy as fw
from fuzzywuzzy import fuzz
import pyodbc
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from datetime import datetime
from dateutil import parser
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

nameFile = pd.DataFrame()
nameCheckComplete = False

class CredentialDialog:
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.top.title("Enter Credentials")

        tk.Label(self.top, text="Username").pack(pady=5)
        self.username = tk.Entry(self.top)
        self.username.pack(pady=5)

        tk.Label(self.top, text="Password").pack(pady=5)
        self.password = tk.Entry(self.top, show="*")
        self.password.pack(pady=5)

        tk.Button(self.top, text="OK", command=self.ok).pack(pady=5)

        self.top.transient(parent)
        self.top.grab_set()
        parent.wait_window(self.top)

    def ok(self):
        self.user = self.username.get()
        self.passwd = self.password.get()
        self.top.destroy()

def get_credentials():
    cred_dialog = CredentialDialog(root)
    return cred_dialog.user, cred_dialog.passwd

def nameCheck():
    """
    Function to input name file and store for future use.
    """
    global nameFile, nameCheckComplete 

    userFileInput = filedialog.askopenfilename(title="Select name file", filetypes=[("Excel Files", "*.xlsx")])
    if userFileInput:
        tempFrame = pd.read_excel(userFileInput) 
        nameColumn = tempFrame.columns[0]
        
        if len(tempFrame.columns) > 1:
            idColumn = tempFrame.columns[1]
        else:
            idColumn = None

        if isinstance(tempFrame[nameColumn].iloc[0], str):
            namesToCheck = tempFrame[nameColumn].apply(lambda x: x.strip().lower())

            if idColumn and pd.api.types.is_numeric_dtype(tempFrame[idColumn]):
                idToCheck = tempFrame[idColumn].apply(lambda x: str(int(x)).strip() if not pd.isna(x) else '')
                nameFile = pd.DataFrame({'names': namesToCheck, 'ids': idToCheck})
                messagebox.showinfo("", "Name and ID file successfully stored.")
            else:
                nameFile = pd.DataFrame({'names': namesToCheck})
                messagebox.showinfo("", "Name file successfully stored.")
            
            nameCheckComplete = True
            print(nameFile)
        else:
            messagebox.showerror("Error processing Names file columns.")
    else:
        messagebox.showerror("Error selecting Names file.")

csvMaster = pd.DataFrame()
csvConsolidateComplete = False

def csvConsolidate():
    """
    Function to consolidate csv files into master Excel file.
    """
    global csvMaster, csvConsolidateComplete, nameFile 

    def checkPartialMatch(wcName, threshold=90):
        for name in nameFile['names']:
            similarityThreshold = fuzz.token_sort_ratio(wcName, name)
            if similarityThreshold >= threshold:
                return True
        return False

    userFileInputTwo = filedialog.askopenfilenames(title="Select CSV files", filetypes=[("Comma Delineated Values Files", "*.csv")])
    print(f"Selected files: {userFileInputTwo}")
    if userFileInputTwo:
        csvFrame = [] 
        for i in userFileInputTwo: 
            print(f"Processing file: {i}")
            wcFile = pd.read_csv(i) 
            wcNameColumnMain = str(wcFile.columns[0]) 
            wcNameColumnAlias = str(wcFile.columns[2])
            if isinstance(wcFile[wcFile.columns[0]].iloc[0], str): 
                wcFile["convertedNames"] = wcFile[wcNameColumnMain].apply(lambda x: x.strip().lower()) 
                matchedNames = wcFile[wcFile["convertedNames"].apply(checkPartialMatch)] 
                if not matchedNames.empty: 
                    csvFrame.append(matchedNames)
                    print(f"File added to csvFrame: {i}")
            elif isinstance(wcFile[wcFile.columns[2]].iloc[0], str): 
                wcFile["convertedAlias"] = wcFile[wcNameColumnAlias].apply(lambda x: x.strip().lower())
                matchedAlias = wcFile[wcFile["convertedAlias"].apply(checkPartialMatch)]
                if not matchedAlias.empty:
                    csvFrame.append(matchedAlias)
                    print(f"File added to csvFrame: {i}")
            else:
                 messagebox.showerror("Error checking names threshold.")
        if csvFrame:
            csvMaster = pd.concat(csvFrame, ignore_index=True) 
            messagebox.showinfo("", "CSV files successfully stored.") 
            exportedMasterFile = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]) 
            if exportedMasterFile:
                csvMaster.to_excel(exportedMasterFile, index=False)
                messagebox.showinfo("", "CSV files concatenated and saved to a new Excel file.")
            csvConsolidateComplete = True
        else:
            messagebox.showinfo("", "No matching names found in CSV files.")
    else:
        messagebox.showerror("Error consolidating CSV files.")
    print(csvMaster)

def patronMaster():
    if nameCheckComplete and csvConsolidateComplete: 
        promptWindow = tk.Toplevel(root) 
        promptWindow.title("Select Action.") 

        manualButton = tk.Button(promptWindow, text="Manual File Load", command=manualPatronFile)
        manualButton.pack(pady=10, padx=10)

        webScrapingButton = tk.Button(promptWindow, text="Search World Check", command=webScraping)
        webScrapingButton.pack(pady=10, padx=10)

        sqlButton = tk.Button(promptWindow, text="SQL Operation", command=sqlPatronFile) 
        sqlButton.pack(pady=10, padx=10)

    else:
        messagebox.showerror("Error", "Please complete Name file and CSV file functions before Patron function.")

def checkPatronMatch(row, csvMaster, nameFile, name_threshold=80):
    """
    Checking patron with World Check profile.
    """
    patron_name = row['concatName']
    patron_year = row['YearofBirth']
    patron_country = row['CountryofOriginPatron']
    patron_id = row['idPatron'] if 'idPatron' in row else None

    # Create dictionary to associate names with their IDs
    name_id_dict = dict(zip(nameFile['names'], nameFile['ids']))

    master_names = csvMaster['convertedNames'].tolist()
    master_years = csvMaster['DateofBirthExtracted'].tolist()
    master_country = csvMaster['CountryofOrigin'].tolist()
    master_datasets = csvMaster['Dataset'].tolist()

    for master_name, master_year, master_country, master_dataset in zip(master_names, master_years, master_country, master_datasets):
        name_match_score = fuzz.token_sort_ratio(patron_name, master_name)
        master_id = name_id_dict.get(master_name, None)
        if patron_id and master_id and patron_id == master_id:
            if patron_year == master_year and 'PEP' in master_dataset:
                return True
        elif name_match_score >= name_threshold:
            if patron_year == master_year and patron_country == master_country and 'PEP' in master_dataset:
                return True
    return False

def extract_year_from_dob(dob):
    """
    Transforming date of birth format for analysis.
    """
    if isinstance(dob, str):
        try:
            date_obj = parser.parse(dob, default=datetime(1900, 1, 1))
            if date_obj.year > datetime.now().year:
                date_obj = date_obj.replace(year=date_obj.year - 100)
            return date_obj.strftime('%Y-%m')
        except ValueError:
            pass
        # Fallback patterns if parsing fails
        patterns = [
            '%d-%b-%Y', '%d %b %Y', '%d-%b-%y', '%d %b %y', 
            '%b-%Y', '%b %Y', '%b-%y', '%b %y',
            '%Y-%b-%d', '%Y %b %d'
        ]
        for pattern in patterns:
            try:
                date_obj = datetime.strptime(dob, pattern)
                if date_obj.year > datetime.now().year:
                    date_obj = date_obj.replace(year=date_obj.year - 100)
                return date_obj.strftime('%Y-%m')
            except ValueError:
                continue
    return None

def countryConversion(CountryDescription):
    """
    Transforming country name to abbreviated name (countries represented are from development areas)
    """
    abbrv = {
        'United States of America': 'USA',
        'United Kingdom': 'GBR',
        'Canada': 'CAN',
        'Australia': 'AUS',
        'New Zealand': 'NZL',
        'Great Britain': 'GBR',
        'China': 'CHN',
        'Saudi Arabia': 'SAU',
        'India': 'IND',
        'Iraq': 'IRQ',
        'Taiwan': 'TWN',
        'Thailand': 'THA',
        'Hong Kong': 'HKG'
    }

    if isinstance(CountryDescription, list):
        return [abbrv.get(country, country) for country in CountryDescription]
    else:
        return abbrv.get(CountryDescription, CountryDescription)

def cityExtract(city):
    if isinstance(city, str):
        parts = city.split(',')
        if parts:
            return parts[0].strip().lower()
    return None

def manualPatronFile():
    global csvMaster
    threshold = 90
    patronDataInput = filedialog.askopenfilename(title="Select Patron Excel File", filetypes=[("Excel Files", "*.xlsx")])
    if patronDataInput:
        patronFrame = pd.DataFrame()
        patronFile = pd.read_excel(patronDataInput)
        patronFile["concatName"] = (patronFile.iloc[:,2] + " " + patronFile.iloc[:,4]).str.strip().str.lower()
        patronFile["concatName"] = patronFile["concatName"].astype(str)
        print(f"Patron File:", patronFile)
        matchedPatron = patronFile[patronFile.apply(checkPatronMatch, axis=1, args=(csvMaster, nameFile))] 
        if not matchedPatron.empty:
            patronFrame = matchedPatron
    if not patronFrame.empty:
        exportedPatronMatchFile = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if exportedPatronMatchFile:
            patronFrame.to_excel(exportedPatronMatchFile, index=False)
            messagebox.showinfo("", "Patron matches saved to new Excel file.")
    else:
        messagebox.showerror("", "No Patron matches with world check results.")

def sqlPatronFile(): 
    """
    Function to run SQL scripts and store information.
    """
    global nameFile, csvMaster
    threshold = 90
    server = '???????' # Server hidden for security
    database = '???????' # Server hidden for security
    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                            'SERVER=' + server + ';'
                            'DATABASE=' + database + ';'
                            'Trusted_Connection=yes;')
    print(cnxn)
    cursor = cnxn.cursor()

    matched_ids = pd.DataFrame()
    matched_names = pd.DataFrame()

    if 'ids' in nameFile:
        ids = nameFile['ids'][nameFile['ids'] != ''].tolist()
        if ids:
            idsPlaceholder = ', '.join('?' for _ in ids)
            query = f"""
            SELECT 
            P.[PlayerId]
            , P.[Status]
            , [FirstName]
            , [MiddleName]
            , [LastName]
            , CAST ([Birthday] AS DATE) as "DateofBirth"
            , [CompanyName]
            , [JobTitle]
            , B.[Description] as BusinessTypeDescription
            , [ChristianizedName]
            , A.[Description] as AddressType, PA.[Line1]
            , PA.[Line2]
            , PA.[ZipCode]
            , PA.[ZipPlus]
            , PA.[City]
            , PA.[State]
            , C.[CountryDescription]
            , PA.[Suburb]
            FROM [PlayerManagementViews].[dbo].[Player] P
            LEFT JOIN [PlayerManagementViews].[dbo].[PlayerAddress] PA ON P.PlayerID = PA.PlayerID
            LEFT JOIN [PlayerManagementViews].[dbo].[AddressType] A ON PA.TypeID = A.TypeID
            LEFT JOIN [PlayerManagementViews].[dbo].[Country] C ON PA.CountryID = C.CountryID
            LEFT JOIN [PlayerManagementViews].[dbo].[BusinessType] B ON P.BusinessTypeID = B.BusinessTypeID
            WHERE P.[PlayerId] IN ({idsPlaceholder})"""
            print(query)
            cursor.execute(query, ids)
            rows = cursor.fetchall()
            matched_ids = pd.DataFrame([tuple(row) for row in rows], columns=[x[0] for x in cursor.description])

    firstName = [name.split()[0] for name in nameFile['names'] if len(name.split()) > 0]
    lastName = [name.split()[-1] for name in nameFile['names'] if len(name.split()) > 0]
    firstNamePlaceholder = ', '.join('?' for _ in firstName)
    lastNamePlaceholder = ', '.join('?' for _ in lastName)
    print(firstNamePlaceholder[0])
    print(lastNamePlaceholder[0])
    query = f"""
    SELECT 
    P.[PlayerId]
    , P.[Status]
    , [FirstName]
    , [MiddleName]
    , [LastName]
    , CAST ([Birthday] AS DATE) as "DateofBirth"
    , [CompanyName]
    , [JobTitle]
    , B.[Description] as BusinessTypeDescription
    , [ChristianizedName]
    , A.[Description] as AddressType, PA.[Line1]
    , PA.[Line2]
    , PA.[ZipCode]
    , PA.[ZipPlus]
    , PA.[City]
    , PA.[State]
    , C.[CountryDescription]
    , PA.[Suburb]
    FROM [PlayerManagementViews].[dbo].[Player] P
    LEFT JOIN [PlayerManagementViews].[dbo].[PlayerAddress] PA ON P.PlayerID = PA.PlayerID
    LEFT JOIN [PlayerManagementViews].[dbo].[AddressType] A ON PA.TypeID = A.TypeID
    LEFT JOIN [PlayerManagementViews].[dbo].[Country] C ON PA.CountryID = C.CountryID
    LEFT JOIN [PlayerManagementViews].[dbo].[BusinessType] B ON P.BusinessTypeID = B.BusinessTypeID
    WHERE [FirstName] IN ({firstNamePlaceholder}) AND [LastName] IN ({lastNamePlaceholder})"""
    print(query)
    cursor.execute(query, firstName + lastName)
    rows = cursor.fetchall()
    matched_names = pd.DataFrame([tuple(row) for row in rows], columns=[x[0] for x in cursor.description])

    patronDF = pd.concat([matched_ids, matched_names]).drop_duplicates().reset_index(drop=True)
    
    print(patronDF)
    if not patronDF.empty:
        messagebox.showinfo("", "SQL query successful.")
        patronDF["concatName"] = (patronDF["FirstName"] + " " + patronDF["LastName"]).str.strip().str.lower()
        patronDF["concatName"] = patronDF["concatName"].astype(str)
        patronDF["idPatron"] = patronDF["PlayerId"].astype(str)
        patronDF["DateofBirth"] = pd.to_datetime(patronDF['DateofBirth'], errors='coerce')
        patronDF['YearofBirth'] = patronDF['DateofBirth'].dt.to_period('M').astype(str)
        csvMaster['DateofBirthExtracted'] = csvMaster['Date of Birth'].apply(lambda x: extract_year_from_dob(x))
        csvMaster['DateofBirthExtracted'] = csvMaster['DateofBirthExtracted'].astype(str)
        print(csvMaster["DateofBirthExtracted"])
        print(patronDF["YearofBirth"])
        csvMasterNames = csvMaster['convertedNames'].tolist()
        csvMaster['CountryofOrigin'] = csvMaster['Citizenship'].str.strip().str.lower().astype(str)
        patronDF['CountryofOriginPatron'] = patronDF['CountryDescription'].apply(lambda x: countryConversion(x)).str.strip().str.lower()
        csvMaster['City'] = csvMaster['Place of Birth'].apply(lambda x: cityExtract(x)).str.strip().str.lower()
        patronDF['CityPatron'] = patronDF['City'].str.strip().str.lower()
        if 'ids' in nameFile:
            print(nameFile['ids'])
        print(patronDF['idPatron'])
        print(csvMaster['CountryofOrigin'])
        print(patronDF['CountryofOriginPatron'])
        print(csvMaster['City'])
        print(patronDF['CityPatron'])
        patronDF['MatchFound'] = patronDF.apply(lambda x: checkPatronMatch(x, csvMaster, nameFile), axis=1)
        matchedPatron = patronDF[patronDF['MatchFound']]
        if not matchedPatron.empty:
            export_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if export_file_path:
                matchedPatron.to_excel(export_file_path, index=False)
                messagebox.showinfo("", "Matched Patron records exported.")
            else:
                messagebox.showerror("", "No file path provided for export.")
        else:
            messagebox.showinfo("", "No matches found between patron records and world check results.")
    else:
        messagebox.showinfo("", "No Patron records found in the database.")

def webScraping():
    global nameFile

    username, password = get_credentials()
    if not username or not password:
        messagebox.showerror("Error", "Username or Password not provided.")
        return

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--log-level=0')  # Enable verbose logging

    driver = webdriver.Chrome(options=options)

    try:
        driver.get("??????") # Site hidden for security

        # Wait for the username input and enter the username
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "User ID"))).send_keys(username)
        driver.find_element(By.NAME, "password").send_keys(password)
        driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)

        # Ensure we are logged in by waiting for a known element on the dashboard
        first_name = nameFile['names'].iloc[0]

        # Wait for the name input field
        name_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "name")))
        name_input.clear()
        name_input.send_keys(first_name)
        
        # Click the Search button
        search_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='Search']"))
        )
        search_button.click()

        # Wait for the Export button and click it
        export_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Export')]"))
        )
        driver.execute_script("arguments[0].click();", export_button)

        # Confirm the export
        confirm_export_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Export') and not(contains(text(),'Hover'))]"))
        )
        driver.execute_script("arguments[0].click();", confirm_export_button)

        # Click the Back to Screening button
        back_to_screening_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Back to Screening')]"))
        )
        driver.execute_script("arguments[0].click();", back_to_screening_button)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        driver.quit()

root = tk.Tk()
root.title("World Check Patron Consolidator")
root.geometry("600x300")

nameFileButton = tk.Button(root, text="Select Names File", command=nameCheck)
nameFileButton.pack(pady=10, padx=10)

csvConsolidateButton = tk.Button(root, text="Select CSV Files", command=csvConsolidate)
csvConsolidateButton.pack(pady=10, padx=10)

patronMasterButton = tk.Button(root, text="Activate Patron Check", command=patronMaster)
patronMasterButton.pack(pady=10, padx=10)

root.mainloop()
