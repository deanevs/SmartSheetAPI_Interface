from decouple import config

smartsheet_token = config('smartsheet_token')

# these are to be added each year since limited to 20,000 rows
# need to check and give warnings if approaching
CORRECTIVES_2020 = config('CORRECTIVES_2020')
CORRECTIVES_2021 = config('CORRECTIVES_2021')
CORRECTIVES_2022 = config('CORRECTIVES_2022')
CORRECTIVES_2023 = config('CORRECTIVES_2023')
KEYWORD_SHEET = config('KEYWORD_SHEET')


SS_INT_WEEKLY = config('SS_INT_WEEKLY')

SS_IMP_LUP_ORIG = config('SS_IMP_LUP_ORIG')

# Nuvolo Siebel
SS_SIEBEL_CLOSED = config('SS_SIEBEL_CLOSED')
SS_INT_SUMMARY = config('SS_INT_SUMMARY')
SIEBEL_FOLDER_ID = config('SIEBEL_FOLDER_ID')  # NUVOLO_INTERFACE / SIEBEL
NUVOLO_INTERFACE_LOG = config('NUVOLO_INTERFACE_LOG')
NUVOLO_MUST_FOLDER_ID = config('NUVOLO_MUST_FOLDER_ID')  # 8974313631049604

# SmartSheet unique to Siebel
TEMPLATE_SPIRE_SIEBEL_ARCHIVED = config('TEMPLATE_SPIRE_SIEBEL_ARCHIVED')
TEMPLATE_SPIRE_MUST_ARCHIVED = config('TEMPLATE_SPIRE_MUST_ARCHIVED')



# CNL SmartSheet
SH_CNL_ID = config('SH_CNL_ID')

# TBC Future contract costs
SS_TBC_COSTS = config('SS_TBC_COSTS')
SS_TBC_LOG = config('SS_TBC_LOG')

# SMAX CORE
SS_INTERIM_DEBRIEF = config('SS_INTERIM_DEBRIEF')







