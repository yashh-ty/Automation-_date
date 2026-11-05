#
# from cleaning_data import formatting_data
#
# mb_mapper_excel={('Fund Transfer', 'IMPS', ''): ('Fund transfer','IMPS'),
#                  ('Fund Transfer', 'NEFT', ''): ('Fund transfer','NEFT'),
#                  ('Fund Transfer', 'RTGS', ''): ('Fund transfer','RTGS'),
#                  ('Fund Transfer', 'Within bank', ''): ('Fund transfer','Within bank'),
#                  ('Fund Transfer', 'SI-RECCURRING', ''): ("Fund transfer","SI RECURRING"),
#                  ('Fund Transfer', 'SI-ONE TIME', ''): ('Fund transfer','SI-ONE TIME'),
#                  ('Fund Transfer', 'Self Transfer', ''): ('Fund transfer','Self Transfer'),
#                  ('UPI Contribution', 'Merchant', 'BOI-Accounts'): ('UPI','To Merchant BOIAccount'),
#                  ('UPI Contribution', 'Merchant', 'Non-BOI Accounts'): ("UPI","To Merchant Non-BOI Account"),
#                  ('UPI Contribution', 'Retail', 'BOI-Accounts'): ('UPI','To Retail BOI Account'),
#                  ('UPI Contribution', 'Retail', 'Non-BOI Accounts'): ('UPI','To Retail Non-BOI Account'),
#                  ('UPI VPA Creation', '', ''): ('UPI','VPA Creation'),
#                  ('Credit card bill payment', '', ''): ('Credit-card Payment','CC-Payment'),
#                  ('SI Executed', '', 'ACTIVE'): ('SI Executed','ACTIVE'),
#                  ('SI Executed', '', 'COMPLETED'): ('SI Executed','COMPLETED'),
#                  ('SI Executed', '', 'CANCELLED'): ('SI Executed','CANCELLED'),
#                  ('SI Executed', '', 'PAUSED'): ('SI Executed','PAUSED'),
#                  ('Bill payments', '', ''): ('BillPay BillDesk','BPAY'),
#                  ('FD contribution', '', ''): ('TD Contribution','FD'),
#                  ('RD contribution', '', ''): ('TD Contribution','RD'),
#                  ('PPF contribution', '', ''): ('PPF Contribution','PPF'),
#                  ('SSY contribution', '', ''): ('SSY Contribution','SSY'),
#                  ('NPS contribution', 'Kfin', ''): ('NPS-KFINTECH Contribution','Kfintech trnx'),
#                  ('NPS contribution', 'NSDL', ''): ('NPS-NSDL Contribution','NSDL Contribution'),
#                  ('NPS contribution', 'NSDL SI recurring', ''): ('NPS-NSDL Contribution','NSDL-SI Recurring'),
#                  ('SGB contribution', '', ''): ('SGB Contribution','SGB'),
#                  ('FRSB contribution', '', ''): ('FRSB Contribution','FRSB'),
#                  ('MF lumpsum contribution', '', ''): ('MF-LUMSUM','MF-LUMPSUM'),
#                  ('BOI Multi Asset Allocation Fund(NFO)', '', ''): ('Multi Asset Allocation Fund','NFO'),
#                  ('APY Contribution', '', ''): None,
#                  ('Mssc Contribution', '', ''): None,
#                  ('PMSBY', '', ''): None,
#                  ('PMJJBY', '', ''): None,
#                  ('MF SIP contribution', '', ''): ('MF XSIP Order','MF-SIP'),
#                  ('Insurance purchases', '', ''): None,
#                  ('ASBA Applications', '', ''): ('ASBA Application','ASBA'),
#                  ('FD closed', '', ''): ('TD Closed','FD-Closed'),
#                  ('RD closed', '', ''): ('TD Closed','RD-Closed'),
#                  ('OD opened against FD', '', ''): ('FD','OD Contribution'),
#                  ('OD opened against RD', '', ''): ('RD','OD Contribution'),
#                  ('PPF', '', ''): ('PPF Opened','PPF-Opened'),
#                  ('SSY', '', ''): ('SSY Opened','SSY-Opened'),
#                  ('NPS', 'KFin', ''): ("NPS-KFINTECH account opened","Kfintech Account Opened"),
#                  ('NPS', 'NSDL', ''): ("NPS-NSDL account opened","NSDL Account Opened"),
#                  ('Insurance', 'Reliance', 'Vehicle'): ('INSUR Reliance Vehicle Premium','Insurance Purchase Contribution'),
#                  ('Insurance', 'Reliance', 'Health'): ('INSUR Reliance Health Premium','Insurance Purchase Contribution'),
#                  ('Insurance', 'SUD', 'Group'): ('SUDLIFE','Insurance Sudlife Purchase Contribution'),
#                  ('Insurance', 'SUD', 'Century Royale'): None,
#                  ('Insurance', 'SUD', 'Posh Sanchay'): None}
#
# def matching(report_mapper_excel,daily_data,monthly_data):
#     final_output={}
#     for i in report_mapper_excel.values():
#             if i==None:
#                 final_output[i] =('0','0','0','0','0','0','0','0')
#             elif i in daily_data and i not in monthly_data:
#                 final_output[i] = daily_data[i] +('0','0','0','0')
#             elif i not in daily_data and i  in monthly_data:
#                 final_output[i] = ('0','0','0','0')+monthly_data[i]
#             else:
#                 final_output[i] = daily_data[i] + monthly_data[i]
#
#
#     return final_output
#
# # data=matching(mb_mapper_excel,formatting_data(r"C:\Users\Alchemy\Desktop\27-09-2026 Report\MB-DATA DAILY"),formatting_data(r"C:\Users\Alchemy\Desktop\27-09-2026 Report\MB-DATA MONTHLY"))
# # print(len(data))
# # for i,j in data.items():
# #     print(i,j)
#
# # def cleaning_data(report_mapper_excel,daily_data,monthly_data):
# #     final_output={}
# #     for i in report_mapper_excel.values():
# #         if i in daily_data:
#         # if i in daily_data and monthly_data:
#         #     final_output[i] = daily_data[i] + monthly_data[i]
#         # else:
#         #     final_output[i] = ('0','0','0','0')
#
#
#












