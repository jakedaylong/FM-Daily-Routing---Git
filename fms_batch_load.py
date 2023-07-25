'''Main file for creating batch load for MS FMS system'''
import os
import datetime
import pandas as pd
import fms_batch_modules as fbm

def main():
    '''Main function for FMS Batch Load xlsx generation'''
    script_dir = os.getcwd()
    driver_dir = os.path.join(script_dir, 'driver_info/')
    route_import = pd.read_excel(driver_dir + 'driver_routes.xlsx')
    vehicle_assignment = pd.read_excel(driver_dir + 'driver_vehicle_assignment.xlsx')
    time_delta = int(input('Please enter the number of days in advance you would like to create this file for: '))
    ref_date = datetime.datetime.now()
    ref_date += datetime.timedelta(days=time_delta)
    ref_day = ref_date.strftime("%d")
    ref_month = ref_date.strftime("%m")
    ref_year = ref_date.strftime("%Y")
    pickup_bldg = 123
    pickup_room = 0
    route_inc = 1
    batch_load_df = fbm.create_empty_df()
    batch_load_df = fbm.batch_load_df_creation(
        vehicle_assignment,
        ref_day,
        ref_month,
        ref_year,
        pickup_bldg,
        pickup_room,
        route_import,
        route_inc,
        batch_load_df)
    fbm.batch_load_export(ref_year, ref_month, ref_day, batch_load_df)

if __name__ == "__main__":
    main()
