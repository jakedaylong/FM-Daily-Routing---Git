'''module file for creating batch load for MS FMS system'''
# pylint: disable=abstract-class-instantiated
# pylint: disable=line-too-long
import pandas as pd

def batch_load_df_creation(vehicle_assignment, ref_day, ref_month, ref_year, pickup_bldg, pickup_room, route_import, route_inc, batch_load_df):
    '''Creates and populated dataframe from known values and input values'''
    for row in route_import.iterrows():
        if row[1]['Start'] == "PM":
            pickup_datetime = str(ref_month)+"/"+str(ref_day)+"/"+str(ref_year)+" "+ '12:30:00PM'
            delivery_datetime = str(ref_month)+"/"+str(ref_day)+"/"+str(ref_year)+" "+ '3:00:00PM'
        else:
            pickup_datetime = str(ref_month)+"/"+str(ref_day)+"/"+str(ref_year)+" "+ '8:00:00AM'
            delivery_datetime = str(ref_month)+"/"+str(ref_day)+"/"+str(ref_year)+" "+ '11:00:00AM'
        driver_vehicle = vehicle_assignment.get(vehicle_assignment['driver'] == route_import.loc[row[0]]['driver'])
        append_dict = [{'order_type':'Distro Delivery',
                        'reference_no':'mail-'+ str(ref_month) + str(ref_day) +'-'+ str(route_inc),
                        'pickup_bldg':pickup_bldg,
                        'pickup_room':pickup_room,
                        'pickup_datetime': pickup_datetime,
                        'delivery_bldg': row[1]['BLDG'],
                        'delivery_room': 0,
                        'delivery_datetime':delivery_datetime,
                        'group': 'group',
                        'team':'Distro',
                        'delivery_vehicle_type':'Van',
                        'container_type':'Cart|1',
                        'fixed_route':'yes',
                        'vehicle':driver_vehicle.iloc[0][1]}]
        batch_load_df = batch_load_df.append(append_dict, ignore_index=True)
        route_inc += 1
    return batch_load_df

def batch_load_export(ref_year, ref_month, ref_day, batch_load_df):
    '''writes final xlsx file from dataframe created within batch_load_df_creation module'''
    batch_load_filename = 'FMS_ORDERS_group'+str(ref_year)+str(ref_month)+str(ref_day)+'.xlsx'
    writer = pd.ExcelWriter(batch_load_filename, engine='xlsxwriter')
    batch_load_df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


def create_empty_df():
    '''Creates empty dataframe with appropriate column headers for import'''
    batch_load_df = pd.DataFrame(columns=['order_type',
                                    'reference_no', 
                                    'pickup_bldg', 
                                    'pickup_room', 
                                    'pickup_datetime', 
                                    'delivery_bldg', 
                                    'delivery_room', 
                                    'delivery_datetime', 
                                    'event_time', 
                                    'group', 
                                    'team', 
                                    'notes', 
                                    'delivery_vehicle_type', 
                                    'container_type', 
                                    'fixed_route', 
                                    'vehicle'])
    return batch_load_df
