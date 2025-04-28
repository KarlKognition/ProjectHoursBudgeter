from os import path
from abc import ABC, abstractmethod
import yaml
import phb_app.wizard.constants.ui_strings as st

class YamlHandler(ABC):
    '''Abstract class for deserialising yaml files.'''
    CONFIG_PATH = path.join(path.dirname(__file__), "config_data.yaml")

    def __init__(self):
        super().__init__()
        self.encoding = st.SpecialStrings.UTF_8.value
        self._load_yaml_data()
    
    @abstractmethod
    def _process_yaml(self,
                      yaml_data: dict) -> None:
        '''Abstract method. Processes the data of the yaml file.'''
        raise NotImplementedError # To override

    def _load_yaml_data(self) -> None:
        '''Template for loading yaml data.'''
        try:
            with open(self.CONFIG_PATH, 'r', encoding=st.SpecialStrings.UTF_8.value) as yaml_file:
                yaml_data = yaml.safe_load(yaml_file)
                self._process_yaml(yaml_data)
        except FileNotFoundError as exc:
            raise FileNotFoundError("The config file could not be found.") from exc
        except UnicodeDecodeError as exc:
            raise exc
