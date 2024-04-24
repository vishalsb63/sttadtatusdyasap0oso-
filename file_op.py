import pandas as pd
from openpyxl import load_workbook
import numpy as np

def file_operations(csv_file_path, workbook_path, output_dir):
    try:
        
        data = pd.read_csv(csv_file_path, header=0, low_memory=False)
        data2=data.copy()

        data1=pd.read_excel(workbook_path)

        filtered_data = data[data['TOPECO'].isin(['_Party', '0_Party'])]
        filtered_data = filtered_data[filtered_data['SALERID'].isnull()]
        filtered_data = filtered_data[filtered_data['CUSTOMERTYPE'] == 'noncustomer']
        filtered_data['Counterparties Type'] = np.where(filtered_data['COUNTERP_TRADEPARTYID'].isin(data1['COUNTERP_TRADEPARTYID']),'9 Large Banks without Sales id','Exchange_Broker')

        filtered_data2 = data[data['TOPECO'].isin(['0_Portfolio', '_Party'])]
        
        #pivot_table_1=pd.pivot_table(data2,index='INSTRUMENT_TYPE',values='COUNTERP_TRADEPARTYID', aggfunc=lambda x: len(set(x)) if)

        filtered_data_1 = data2[data2['TOPECO'].isin(['1_Portfolio', '1_Party'])]

        
        filtered_data_1 = filtered_data_1[filtered_data_1['CUSTOMERTYPE'] == 'customer']
        filtered_data_1['Counterparties Type'] = np.where(filtered_data_1['COUNTERP_TRADEPARTYID'].isin(data1['COUNTERP_TRADEPARTYID']),'9 Large Banks without Sales id','Exchange_Broker')

        filtered_csv = f"{output_dir}/{csv_file_path.split('/')[-1].replace('.csv', '_filtered.csv')}"
        filtered_data.to_csv(filtered_csv, index=False)
        print(f"Filtered data saved to {filtered_csv}")
        print(filtered_data.head())

        filtered_csv1 = f"{output_dir}/{csv_file_path.split('/')[-1].replace('.csv', 'cus_filtered.csv')}"
        filtered_data_1.to_csv(filtered_csv1, index=False)
        print(f"Filtered data saved to {filtered_csv1}")
        print(filtered_data_1.head())


        filtered_data2 = data[data['TOPECO'].isin(['0_Party', '_Party'])]

        pivot_table = pd.pivot_table(filtered_data2, 
                             values=['COUNTERP_TRADEPARTYID', 'ABS_EXCHANGEDAMOUNT_USD'],  
                             index='INSTRUMENT_TYPE',  
                             columns='CUSTOMERTYPE',  
                             aggfunc={'COUNTERP_TRADEPARTYID': 'count', 'ABS_EXCHANGEDAMOUNT_USD': 'sum'},  
                             fill_value=0)
        

        pivot_table['Total_trades'] = pivot_table[('COUNTERP_TRADEPARTYID', 'customer')] + pivot_table[('COUNTERP_TRADEPARTYID', 'noncustomer')]
        pivot_table['Total_amount'] = pivot_table[('ABS_EXCHANGEDAMOUNT_USD', 'customer')] + pivot_table[('ABS_EXCHANGEDAMOUNT_USD', 'noncustomer')]

        vm7_data = pd.DataFrame({'count': (pivot_table[('COUNTERP_TRADEPARTYID', 'customer')]/ pivot_table['Total_trades']) * 1,'value': (pivot_table[('ABS_EXCHANGEDAMOUNT_USD', 'customer')]  / pivot_table['Total_amount']) * 1})
        
        total_trades_cus=pivot_table[('COUNTERP_TRADEPARTYID', 'customer')].sum()
        total_trades=pivot_table['Total_trades'].sum()
        total_amount_cus=pivot_table[('ABS_EXCHANGEDAMOUNT_USD', 'customer')].sum()
        total_amount=pivot_table['Total_amount'].sum()
        
        vm7_data.loc['Grand Total'] ={'count':(total_trades_cus/total_trades)*1,'value':(total_amount_cus/total_amount)*1}
        
        pivot_table.loc['Grand Total'] = pivot_table.sum()

        

        

        pivot_table.columns = [f'{col[0]}_{col[1]}' for col in pivot_table.columns]

        filtered_noncustomer = data[data['TOPECO'].isin(['0_Party', '_Party'])]
        filtered_noncustomer = filtered_noncustomer[(filtered_noncustomer['CUSTOMERTYPE'] == 'noncustomer') & (filtered_noncustomer['INSTRUMENT_TYPE'].isin(['Other_Derivatives']))]



        #table2
        pivot_table1 = pd.pivot_table(
                            filtered_noncustomer,
                            index='PRODUCT_STRUCTURE_TYPE',
                            values=['ABS_EXCHANGEDAMOUNT_USD', 'COUNTERP_TRADEPARTYID'],
                            aggfunc={'ABS_EXCHANGEDAMOUNT_USD': 'sum', 'COUNTERP_TRADEPARTYID': 'count'},
                            fill_value=0,
                            margins=True,
                            margins_name='Grand Total')
        
        vm7_data1 = pd.DataFrame({'count': (pivot_table1['COUNTERP_TRADEPARTYID'])/ (3966) * 1,'value': (pivot_table1['ABS_EXCHANGEDAMOUNT_USD'])  / (18777438375.3192) * 1})
    
        
        

        

        with pd.ExcelWriter(output_dir+'/output_pivot_table.xlsx') as writer:
            pivot_table.to_excel(writer, sheet_name='Pivot Table')
            workbook  = writer.book
            worksheet = writer.sheets['Pivot Table']
            vm7_start_col = len(pivot_table.columns) + 2 
            vm7_start_col1 = len(pivot_table1.columns) + 2 
            start_row = 10 + len(pivot_table.index) + 10

            

            

            

            vm7_data.to_excel(writer, sheet_name='Pivot Table', startrow=1, startcol=vm7_start_col)
            vm7_header_format = workbook.add_format({'bold': True, 'fg_color': '#ADD8E6', 'align': 'center'})  

            vm7_data1.to_excel(writer, sheet_name='Pivot Table', startrow=start_row+1, startcol=vm7_start_col1)
            
            percentage_format = workbook.add_format({'num_format': '0.00%'})

            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'bg_color': '#ADD8E6',
                'border': 1
            })

            number_format = workbook.add_format({'num_format': '#,##0'})

        
        
            worksheet.merge_range(0, vm7_start_col, 0, vm7_start_col + 2, 'vm7', vm7_header_format)
            worksheet.write(1, vm7_start_col, 'Instrument Type', workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))
            worksheet.write(1, vm7_start_col+1, 'count', workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))
            worksheet.write(1, vm7_start_col + 2, 'value', workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))
            
            worksheet.merge_range(start_row-1, vm7_start_col1, start_row, vm7_start_col1 + 2, 'vm7', vm7_header_format)
            worksheet.write(start_row, vm7_start_col1, 'Instrument Type', workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))
            worksheet.write(start_row, vm7_start_col1+1, 'count', workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))
            worksheet.write(start_row, vm7_start_col1 + 2, 'value', workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))

            
            

           
            for col_num, value in enumerate(pivot_table.columns.values):
                worksheet.write(0, col_num+1, value, header_format)
                worksheet.set_column(col_num, col_num+1, 18, number_format)

           
            grand_total_row = len(pivot_table) + 2  
            worksheet.write('A' + str(grand_total_row), 'Grand Total', workbook.add_format({'bold': True,'bg_color':'#ADD8E6'}))
            for i in range(1, len(pivot_table.columns) + 1):
                worksheet.write(grand_total_row - 1, i, pivot_table.iloc[-1, i - 1], workbook.add_format({'bold': True, 'num_format': '#,##0','bg_color':'#ADD8E6'}))

            
            for col_num in range(2):
                worksheet.set_column(vm7_start_col + col_num, vm7_start_col + col_num, 18, percentage_format)
                worksheet.set_column(vm7_start_col + col_num+1, vm7_start_col + col_num+1, 18, percentage_format)

            for col_num in range(2):
                worksheet.set_column(vm7_start_col1 + col_num, vm7_start_col1 + col_num, 18, percentage_format)
                worksheet.set_column(vm7_start_col1 + col_num+1, vm7_start_col1 + col_num+1, 18, percentage_format)

                

            
            
            pivot_table1.to_excel(writer, sheet_name='Pivot Table',startrow=start_row)

            for col_num, value in enumerate(pivot_table1.columns.values):
                worksheet.write(0, col_num+1, value, header_format)
                worksheet.set_column(col_num, col_num+1, 18, number_format)


            

           
                
    
        
        
    except Exception as e:
        print(f"Error processing the CSV file: {e}")


file_operations("C:/Users/vbabu021524/OneDrive - GROUP DIGITAL WORKPLACE/Desktop/9banks_reports and inputs/RU151_Numerator-21-NOV-22.csv", "C:/Users/vbabu021524/OneDrive - GROUP DIGITAL WORKPLACE/Desktop/9banks_reports and inputs/9large_banks_List.xlsx", "C:/Users/vbabu021524/OneDrive - GROUP DIGITAL WORKPLACE/Desktop/9banks_reports and inputs")
