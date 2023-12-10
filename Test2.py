import ntplib
from time import ctime
from datetime import datetime, date

EXP_DAY = '2024-01-03'

def print_time():

    try:
        ntp_client = ntplib.NTPClient()
        response = ntp_client.request('pool.ntp.org')
        # print(ctime(response.tx_time))
        now_date = datetime.fromtimestamp(response.tx_time).strftime("%Y-%m-%d")
    except Exception as ex:
        print("ERROR: Cannot get internet date!!!")

    exp_date_formatted = datetime.strptime(EXP_DAY, "%Y-%m-%d").date()

    #if exp_date_formatted >= now_date:


if __name__ == '__main__':
    print_time()