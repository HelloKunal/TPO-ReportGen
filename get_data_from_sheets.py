from __future__ import print_function
from operator import itemgetter
import os.path
import statistics

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

BRANCH_DETAILS = {
    "CSE": ["Computer Science and Engineering", 'I', 4],
    "ECE": ["Electronics and Communications Engineering", 'H', 5],
    "EE": ["Electrical Engineering", 'F', 6],
    "ME": ["Mechanical Engineering", 'G', 7],
    "CHEM": ["Chemical Engineering", 'K', 8],
    "CE": ["Civil Engineering", 'E', 9],
    "MME": ["Metallurgy and Material Science Engineering", 'J', 10],
}
# The ID and range of a sample spreadsheet.
class GetValues:
    def __init__(self, branch) -> None:
        creds = None
        self.branch = branch
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'static/credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        self.creds = creds

    def compl_proc(self):
        self.company_values()
        super_list = self.student_values()
        super_list.append(self.special_values())
        super_list.append(self.branch_companies_with_selections)
        super_list.append(self.branch_companies_without_selections)
        super_list.append(self.companies_with_selections)
        # print(super_list)
        return super_list

    def student_values(self):
        
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            # Call the Sheets API
            
            SPREADSHEET_ID = '19orX5CPrQ7GPyZvQKOHjVZrhKK2dfMLFZj4B7YbFu54'
            RANGES = ['BTech 6M+FTE/PPO!A2:C2000', 'BTech 6M+FTE/PPO!D2:D2000', 'BTech 6M+FTE/PPO!E2:E2000', 'BTech 6M+FTE/PPO!F2:F2000']
            VALUE_RENDER_OPTION = 'UNFORMATTED_VALUE'
            sheet = service.spreadsheets()
            result = sheet.values().batchGet(spreadsheetId=SPREADSHEET_ID,
                                        ranges=RANGES, valueRenderOption=VALUE_RENDER_OPTION).execute()
            # print(result)
            fresult = []
            for r in result['valueRanges']:
                fresult.append((r['values']))
            # print(fresult)
            data = fresult[0]
            company = fresult[1]
            ppocompany = fresult[2]
            sixmcompany = fresult[3]
            
            # if not values:
            #     print('No data found.')
            #     return

            # print('Name, Major:')
            # print(values)
            fte_data = []
            ppo_data = []
            sixm_data = []
            total_offers = 0
            number_of_placed = 0
            number_of_sixm = 0
            all_packages = []
            for i in range(0, len(data)):
                if len(data[i]) == 0:
                    break
                if data[i][1] == BRANCH_DETAILS[self.branch][0]:
                    if  len(company[i]):
                        total_offers += 1
                        number_of_placed += 1
                        one_data = []
                        one_data.append(data[i][2])
                        one_data.append(data[i][0])
                        company_name = '*'
                        company_name = company[i][0]
                        try:
                            if ',' in company_name:
                                all_offers = []
                                companies = company_name.split(", ")
                                for com in companies:
                                    its_ctc = self.company_dict[com][2]
                                    all_offers.append([its_ctc, com])
                                    sorted_all_offers = sorted(all_offers, reverse=True)
                                    company_name = sorted_all_offers[0][1]
                        except:
                            company_name = company[i][0]
                        one_data.append(company_name)
                        try:
                            one_data.append(self.company_dict[company_name][2])
                            all_packages.append(self.company_dict[company_name][2])
                        except:
                            one_data.append('*')

                        fte_data.append(one_data)
                    elif i < len(ppocompany) and len(ppocompany[i]):
                        total_offers += 1
                        number_of_placed += 1
                        one_data = []
                        one_data.append(data[i][2])
                        one_data.append(data[i][0])
                        company_name = '*'
                        company_name = ppocompany[i][0]
                        if ',' in company_name:
                            all_offers = []
                            companies = company_name.split(", ")
                            for com in companies:
                                its_ctc = self.company_dict[com][2]
                                all_offers.append([its_ctc, com])
                                sorted_all_offers = sorted(all_offers, reverse=True)
                                company_name = sorted_all_offers[0][1]
                        one_data.append(company_name)
                        try:
                            one_data.append(self.company_dict[company_name][2])
                        except:
                            one_data.append('*')

                        ppo_data.append(one_data)
                    if i < len(sixmcompany) and len(sixmcompany[i]):
                        total_offers += 1
                        number_of_sixm += 1
                        one_data = []
                        one_data.append(data[i][2])
                        one_data.append(data[i][0])
                        company_name = '*'
                        company_name = sixmcompany[i][0]
                        if ',' in company_name:
                            all_offers = []
                            companies = company_name.split(", ")
                            for com in companies:
                                its_ctc = self.company_dict[com][2]
                                all_offers.append([its_ctc, com])
                                sorted_all_offers = sorted(all_offers, reverse=True)
                                company_name = sorted_all_offers[0][1]
                        one_data.append(company_name)
                        try:
                            one_data.append(self.company_dict[company_name][2])
                        except:
                            one_data.append('*')

                        sixm_data.append(one_data)


            self.total_offers = total_offers
            self.number_of_placed = number_of_placed
            self.number_of_sixm = number_of_sixm
            for i in all_packages:
                if i == '*':
                    all_packages.remove(i)
            self.all_packages = all_packages

            # print(fte_data)
            fte_data = sorted(fte_data, key=itemgetter(2, 1))
            ppo_data = sorted(ppo_data, key=itemgetter(2, 1))
            sixm_data = sorted(sixm_data, key=itemgetter(2, 1))
            for i in range(1, len(fte_data)+1):
                fte_data[i-1].insert(0, i)
            for i in range(1, len(ppo_data)+1):
                ppo_data[i-1].insert(0, i)
            for i in range(1, len(sixm_data)+1):
                sixm_data[i-1].insert(0, i)
            # print(fte_data)
            # print(ppo_data)
            # print(sixm_data)
            # return data
            return [fte_data, ppo_data, sixm_data]
        except HttpError as err:
            print(err)

    
    def special_values(self):          
        
        service = build('sheets', 'v4', credentials=self.creds)
        SPREADSHEET_ID = '19orX5CPrQ7GPyZvQKOHjVZrhKK2dfMLFZj4B7YbFu54'
        RANGES = ['Sheet1!A5:A2000', 'Eligible Count!E'+str(BRANCH_DETAILS[self.branch][2])]
        VALUE_RENDER_OPTION = 'UNFORMATTED_VALUE'
        sheet = service.spreadsheets()
        result = sheet.values().batchGet(spreadsheetId=SPREADSHEET_ID,
                                    ranges=RANGES, valueRenderOption=VALUE_RENDER_OPTION).execute()
        fresult = []
        # print(result)
        for r in result['valueRanges']:
            fresult.append((r['values']))
        # print(fresult[0])
        special_list = []
        total_companies_count = 0
        while total_companies_count < len(fresult[0]) and len(fresult[0][total_companies_count]):
            total_companies_count += 1
        special_list.append(total_companies_count)


        strength = fresult[1][0][0]
        special_list.append(strength)

        total_offers = self.total_offers
        special_list.append(total_offers)

        number_of_placed = self.number_of_placed
        special_list.append(number_of_placed)

        number_of_sixm = self.number_of_sixm
        special_list.append(number_of_sixm)

        placement_percent = number_of_placed / strength
        placement_percent *= 100
        placement_percent = round(placement_percent, 2)
        placement_percent = str(placement_percent) + '%'
        special_list.append(placement_percent)

        all_package = self.all_packages
        average_package = sum(all_package) / len(all_package)
        average_package = round(average_package, 2)
        special_list.append(average_package)

        median_package = statistics.median(all_package)
        median_package = round(median_package, 2)
        special_list.append(median_package)

        highest_package = max(all_package)
        special_list.append(highest_package)

        lowest_package = min(all_package)
        special_list.append(lowest_package)
        
        return special_list

    def company_values(self):
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            # Call the Sheets API
            
            SPREADSHEET_ID = '19orX5CPrQ7GPyZvQKOHjVZrhKK2dfMLFZj4B7YbFu54'
            RANGES = ['Sheet1!C5:C500', 'Sheet1!'+BRANCH_DETAILS[self.branch][1]+'5:'+BRANCH_DETAILS[self.branch][1]+'500', 'Sheet1!Z5:Z500']
            VALUE_RENDER_OPTION = 'UNFORMATTED_VALUE'
            sheet = service.spreadsheets()
            result = sheet.values().batchGet(spreadsheetId=SPREADSHEET_ID,
                                        ranges=RANGES, valueRenderOption=VALUE_RENDER_OPTION).execute()
            fresult = []
            for r in result['valueRanges']:
                fresult.append((r['values']))
            # print(fresult[0])
            names = fresult[0]
            count = fresult[1]
            ctc = fresult[2]

            
            # if not values:
            #     print('No data found.')
            #     return

            # print('Name, Major:')
            # print(values)
            company_dict = {}
            branch_companies_with_selections = []
            branch_companies_without_selections = []
            companies_with_selections = []
            for i in range(0, len(names)):
                if len(names[i]) == 0:
                    break
                com_name = names[i][0]
                selections = 0
                ctc_value = '*'
                if i < len(ctc) and len(ctc[i]):
                    ctc_value = ctc[i][0]                    
                if i < len(count) and len(count[i]):
                    selections = count[i][0]
                    branch_companies_with_selections.append(com_name)
                    if count[i][0] == 0:
                        branch_companies_without_selections.append(com_name)
                    else:
                        if ctc_value != '*':
                            companies_with_selections.append([com_name, selections, ctc_value])
                company_dict[com_name] = [com_name, selections, ctc_value]
                # print(company_dict[com_name])
                # Print columns A and E, which correspond to indices 0 and 4.
                # print(row[0], row[6], row[23])
                # print(row)

            # print(company_dict)
            self.company_dict = company_dict
            self.branch_companies_with_selections = branch_companies_with_selections
            self.branch_companies_without_selections = branch_companies_without_selections

            companies_with_selections = sorted(companies_with_selections, key=itemgetter(1, 2), reverse=True)
            for i in range(1, len(companies_with_selections)+1):
                companies_with_selections[i-1].insert(0, i)
            self.companies_with_selections = companies_with_selections
        except HttpError as err:
            print(err)


if __name__ == '__main__':
    get_values_object = GetValues()
    get_values_object.compl_proc()