from setting_wiz import *
from map_rac import *
from get_data_callrec import *
import setting_wiz

def main():
    for i in range(3):
        try:
            get_data(date)
            time.sleep(5)
            map_data()
            break
        except Exception as ex:
            raise ex
        
if __name__ == '__main__':
    main()