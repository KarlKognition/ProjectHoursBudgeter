from dataclasses import dataclass
from typing import Optional
import phb_app.wizard.constants.ui_strings as st
import phb_app.data.yaml_handler as yh

@dataclass()
class HoursDeviation(yh.YamlHandler):
    '''Data class to define the deviation thresholds for predicted and accumulated hours.'''
    strong_dev: Optional[float] = None
    weak_dev: Optional[float] = None

    def __post_init__(self):
        yh.YamlHandler.__init__(self)
        self._load_yaml_data()

    def _process_yaml(self, yaml_data) -> None:
        '''Processes the yaml data.'''
        self.__dict__.update(yaml_data.get(st.YamlEnum.DEVIATIONS, {}))