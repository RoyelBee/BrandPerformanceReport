import time
# import pandas._libs.tslibs.base

start_time = time.time()

# # ------------- Generate all Data set ------------------------
import AdditonalFiles.generate_raw_data as data

print('Program Executing... ')
data.sales_achiv_trend_data()
data.seen_rx_data()
data.doctor_call_data()

print('All Raw data created \n')
# # --------------------- Send All Data -----------------------
import AdditonalFiles.send_all_mail as all_mail

all_mail.send_all_report()

# ----------- Send Single RSM Mail ----------------------------
import AdditonalFiles.send_rsm_mail as mail

# # --------------- All Users for GPM Tafsir ------------------
# mail.send_report('CBU', 'rejaul.islam@transcombd.com')  # 'abul.basher@skf.transcombd.com'


print('Time takes = ', round((time.time() - start_time) / 60, 2), 'Min')
