import glpi_api, pandas, configparser

class OfficeKeys:
    """Main class to manage writing into glpi"""

    def __init__(self):
        """Initialize basic variables"""
        # Main excel file that application will use to update/add office keys into glpi
        self.master_list = pandas.read_excel(r"Master\masterfile.xlsx", sheet_name="Sheet1")
        # Main list
        self.main_list = []
        # Config file
        self.config = configparser.ConfigParser()
        self.config.read('settings.ini')
        # Url of api in inventory
        self.URL = self.config['URL']['Glpi_url']
        # API token
        self.APPTOKEN = self.config['TOKENS']['App_token']
        # User token
        self.USERTOKEN = self.config['TOKENS']['User_token']
        # if True copy file
        self.file_exists = False
        # Counter for loop
        self.counter = 0
        # if exists
        self.exist = False
        # Id for office in glpi - add below to settings file
        self.office2021 = self.config['OFFICE ID']['Office_2021']
        self.office2019 = self.config['OFFICE ID']['Office_2019']
        self.office2016 = self.config['OFFICE ID']['Office_2016']
        # Id of manufacturer
        self.man_id = self.config['PUBLISHER']['Microsoft']
        # Id of state
        self.przypisany = self.config['STATUS']['Przypisany']
        self.wolny = self.config['STATUS']['Wolny']

        # Main app flag
        self.running = True
        # Error flag
        self.error_flag = False

    def run_app(self):
        """Function that starts application"""
        self._exceldictlist()
        while self.running:
            self._assign_to_variable(self.main_list[self.counter])
            if self._validate_user(self.user):
                self._user_to_name(self.user)
                self._search_user_glpi(self.last_name, self.first_name)
            self._string_to_number(self.office_name)
            self._assign_id_to_version()
            self._assign_id_to_status()
            self._get_name(self.key)

            self._check_if_exists(self.key)
            if not self.error_flag:
                if self.exist:
                    self._update_to_glpi(self.query_id)
                else:
                    self._add_to_glpi()
            if self.error_flag:
                self._error_flag_triggered()
            self._counter()
            self.exist = False
            self.error_flag = False

    def _error_flag_triggered(self):
        """Print information related to which row triggered an error in master file"""
        print(f'Error triggered in row: {self.counter+2} in the master file\n')

    def _validate_user(self, string):
        """Validate if string exists and is a str and return False/True and assign user to null if False"""
        if type(string) == str:
            if len(string) > 0:
                return True
            else:
                self.id = 0
                return False
        else:
            self.id = 0
            return False

    def _string_to_number(self, name):
        """Extract version from name string"""
        try:
            self.version = ""
            for letters in name:
                if letters.isnumeric():
                    self.version = self.version + letters
        except TypeError:
            # Trigger an error flag
            self.error_flag = True

    def _assign_id_to_version(self):
        """Based on version assign id"""
        self.office = ""
        if self.version == "2021":
            self.office = self.office2021
        elif self.version == "2019":
            self.office = self.office2019
        elif self.version == "2016":
            self.office = self.office2016
        else:
            self.office = ""

    def _assign_id_to_status(self):
        """Based on status assign id"""
        self.status_id = ""
        if self.status.lower() in "przypisany":
            self.status_id = self.przypisany
        elif self.status.lower() in "wolny":
            self.status_id = self.wolny
        else:
            self.status_id = 0

    def _get_name(self, serial_key):
        """Take office version and part of serial number to create unique name"""
        try:
            serial_key = serial_key[0:5]
            self.officekeyname = f'Microsoft Office {self.version}({serial_key})'
        except TypeError:
            # Raise error flag
            self.error_flag = True

    def _update_to_glpi(self, query_id):
        """Function that updates object in glpi"""
        try:
            with glpi_api.connect(self.URL, self.APPTOKEN, self.USERTOKEN) as glpi:
                glpi.update('softwarelicense',
                         {'id': f'{query_id}',
                          'name': f'{self.officekeyname}',
                          'softwares_id': f'{self.office}',
                          'users_id': f'{self.id}',
                          'serial': f'{self.key}',
                          'states_id': f'{self.status_id}',
                          'comment': f'Laptop: {self.laptop}\nKonto: {self.konto}\nData dodania: {self.data}'})
        except glpi_api.GLPIError as err:
            print(str(err))

    def _check_if_exists(self, serial):
        """Check if object exist in database and get it's id."""
        criteria = [{'field': 'serial', 'searchtype': 'contains', 'value': serial}]
        forceddisplay = [2]
        try:
            with glpi_api.connect(self.URL, self.APPTOKEN, self.USERTOKEN) as glpi:
                query = glpi.search('SoftwareLicense', criteria=criteria, forcedisplay=forceddisplay)
                if query:
                    self.exist = True
                    for items in query:
                        self.query_id = items.get('2')
                else:
                    self.exist = False
        except glpi_api.GLPIError as err:
            print(str(err))

    def _add_to_glpi(self):
        """Add license to glpi"""
        try:
            with glpi_api.connect(self.URL, self.APPTOKEN, self.USERTOKEN) as glpi:
                glpi.add('SoftwareLicense',
                         {'name': f'{self.officekeyname}',
                          'users_id': f'{self.id}',
                          'softwares_id': f'{self.office}',
                          'serial': f'{self.key}',
                          'states_id': f'{self.status_id}',
                          'comment': f'Laptop: {self.laptop}\nKonto: {self.konto}\nData dodania: {self.data}',
                          'manufacturers_id': self.man_id})
        except glpi_api.GLPIError as err:
            print(str(err))

    def _user_to_name(self, user):
        """Takes full user string and assigns it to name and surname variables"""
        self.first_name = user.split()[0]
        self.last_name = user.split()[1]

    def _assign_to_variable(self, dict):
        """Takes key-value pairs from dictionary and assigns values to variables"""
        self.office_name = dict.get("Wersja")
        self.key = dict.get("Klucz")
        self.user = dict.get("Użytkownik")
        self.laptop = dict.get("Laptop")
        self.konto = dict.get("Konto")
        self.data = dict.get("Data dodania")
        self.status = dict.get("Status")

    def _counter(self):
        """increase counter for the main loop based on length of main_list"""
        self.counter += 1
        if self.counter >= len(self.main_list):
            self.running = False

    def _search_user_glpi(self, surname, firstname):
        """Connect to glpi,search for user to find their id"""
        # This function should be in a loop that individually checks and assigns user combination of while and for loop
        criteria = [{'field': 'realname', 'searchtype': 'contains', 'value': surname},
                    {'link': 'AND', 'field': 'firstname', 'searchtype': 'contains', 'value': firstname}]
        forceddisplay = [2]
        try:
            with glpi_api.connect(self.URL, self.APPTOKEN, self.USERTOKEN) as glpi:
                query = glpi.search('User', criteria=criteria, forcedisplay=forceddisplay)
                for items in query:
                    self.id = items.get('2')
        except glpi_api.GLPIError as err:
            print(str(err))

    def _exceldictlist(self):
        """Put excel rows and columns into dict. and then into list of dicts"""
        for idx, row in self.master_list.iterrows():
            version = row['Wersja']
            key = row['Klucz']
            user = row['Użytkownik']
            pc = row['Laptop']
            acc = row['Konto']
            data = row['Data dodania']
            status = row['Status']
            dict = {"Wersja": version,
                    "Klucz": key,
                    "Użytkownik": user,
                    "Laptop": pc,
                    "Data dodania": data,
                    "Konto": acc,
                    "Status": status}
            self.main_list.append(dict)

if __name__ == "__main__":
    app = OfficeKeys()
    app.run_app()
