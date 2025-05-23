from dataclasses import dataclass

# Structure of input data
@dataclass
class EntryData:
    folderpath: str
    targetrange: str
    workdate: str
    workername: str
    terminalname: str
