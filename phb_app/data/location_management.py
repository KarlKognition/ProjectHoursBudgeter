'''
Module Name
---------
PHB Wizard Location Management

Author
-------
Karl Goran Antony Zuvela

Description
-----------
Data classes for managing the locale data used in the wizard.
'''
from dataclasses import dataclass, field
import phb_app.data.yaml_handler as yh
import phb_app.wizard.constants.ui_strings as st

@dataclass(slots=True)
class FilterHeaders:
    '''Data class for worksheet header strings used for filtering.
    These data are received from LocaleData.'''
    name: str
    proj_id: str # Project number
    description: str # Project's short description
    hours: str
    date: str

@dataclass(slots=True)
class FilePatternData:
    '''Parent data class for establishing the Excel file's naming.'''
    # German or external input timesheets or output budget file
    file_type: str
    # Regular expresion to filter for file in open file dialog.
    file_patterns: list[str]

@dataclass(slots=True)
class InputLocaleData(FilePatternData):
    '''Child data class for establishing the Excel file's locale details.'''
    country: str # Country name
    exp_sheet_name: str # Expected worksheet name
    filter_headers: FilterHeaders = field(default_factory=dict)

    def __post_init__(self):
        '''Init the filter headers from the data received from the country data dataclass.'''
        self.filter_headers = FilterHeaders(**self.filter_headers)

@dataclass(slots=True)
class CountryData(yh.YamlHandler):
    '''Data class using the LocaleData data class to deserialise the yaml config file.'''
    countries: list[InputLocaleData] = field(default_factory=list)

    def __post_init__(self):
        yh.YamlHandler.__init__(self)
        self._load_yaml_data()

    def _process_yaml(self, yaml_data) -> None:
        '''Processes the yaml data.'''
        country_data = yaml_data.get(st.YamlEnum.COUNTRIES, [])
        self.countries = [InputLocaleData(**locale_data) for locale_data in country_data]

    def get_locale_by_country(self, country: str) -> InputLocaleData:
        '''Returns the locale as per given country name.'''
        return next((locale for locale in self.countries if locale.country.lower() == country.lower()), None)
