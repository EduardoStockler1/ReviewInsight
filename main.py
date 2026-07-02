from processing import dataCleaner
from scrapping import googleMapsScrapping

def main () -> None:
    dataCleaner.main()
    googleMapsScrapping.main()
    
        