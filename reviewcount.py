import pandas as pd
import sys, os

def textfile_to_list(filename):
    lines = open(filename, 'r').readlines()
    
    colleagues = []
    for line in lines:
        colleagues.append(line.strip())
    
    return colleagues


# Raw Data to work with from command line
fo_colleagues = textfile_to_list(sys.argv[1])   # .txt file
hskp_colleagues = textfile_to_list(sys.argv[2]) # .txt file
trustyou_raw = pd.read_excel(sys.argv[3])           # .xlsx file
tripadvisor_raw = pd.read_excel(sys.argv[4])        # .xlsx file

# DataFrames here
tripadvisor = pd.DataFrame(tripadvisor_raw)
trustyou = pd.DataFrame(trustyou_raw)


# find the first index of a name in a row of a dataframe
def find_unique_occurrence(dataset, colleague):
        """ Returns the number of rows that a name appears in """
        count = 0
        for row in range(len(dataset)):
            if dataset.loc[row].to_string().lower().find(f' {colleague.lower()} ') > 0:
                 count += 1
        return count


# find the first index of a name in a row of a dataframe
def process_colleague_list(colleagues_List):
    colleague_tributes = {}
    ty_list = []
    ta_list = []
    count = 0
    for colleague in colleagues_List:
        count += 1
        ty_mentions = find_unique_occurrence(trustyou, colleague)
        ta_mentions = find_unique_occurrence(tripadvisor, colleague)
        ty_list.append(ty_mentions)
        ta_list.append(ta_mentions)
        print(f'parsing {count} of {len(colleagues_List)}')
    colleague_tributes = { "colleagues": colleagues_List, "trustyou": ty_list, "tripadvisor": ta_list }
    return colleague_tributes


# final output for file
def output_to_excel(data, wsheetname):
     if not os.path.exists('mentions'):
        os.makedirs('mentions')
     wname = "mentions/" + wsheetname + ".xlsx"
     print('outputting data to excel...')
     outputdf = pd.DataFrame(data)
     outputdf.to_excel( wname, sheet_name="colleagues mentioned")
     print(f"{wname} created")


def run():
    fo_colleagues_by_mentions = process_colleague_list(fo_colleagues)
    hk_colleagues_by_mentions = process_colleague_list(hskp_colleagues)
    output_to_excel(fo_colleagues_by_mentions, "frontOfficeMentions")
    output_to_excel(hk_colleagues_by_mentions, "housekeepingMentions")



if __name__ == '__main__':
     run()