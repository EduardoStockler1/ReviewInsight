from scrapping.googleMapsScrapping.scrapper import main as google_maps
from processing import dataCleaner

def main () -> None:
    google_maps().main()
    dataCleaner.main()
    
        