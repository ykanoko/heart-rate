import pandas as pd

for i in range(3):
    subject = 'S0' + str(i + 4)
    # subject = 'S' + str(i + 16)

    # date1 = '230927_'
    date1 = '230928_'
    # date1 = '231003_'
    # date1 = '231004_'

    # type1 = '_1'
    # type1 = '_2'
    # type1 = '_3'
    # type1 = '_4'
    type1 = '_5'
    # type1 = '_6'
    
    search_time = '13:21:30'

    input_dir = "C:\\Users\\kanok\\Documents\\大学院\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍結果\\"
    input_file = input_dir + date1 + subject + '.csv'
    output_dir = output_dir = 'C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\心拍切り取りデータ\\'

    df = pd.read_csv(input_file, encoding='shift-jis', skiprows=5)

    df['time'] = pd.to_datetime(df['time'])
    # df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0])
    df = df[df['time'].dt.time >= pd.to_datetime(search_time).time()]

    df = df[['RRI']]
    output_file = output_dir + '2309' + subject + type1 + '.csv'
    df.to_csv(output_file, index=False, header=False)
    print('i', i)