import requests
import os
import xlsxwriter


def write_data(url):
    filename = input('filename: ')
    directory = os.path.expanduser('~/Desktop/')
    workbook = xlsxwriter.Workbook(directory + f'{filename}.xlsx')
    worksheet = workbook.add_worksheet()
    colTitles = ['transactionHash', 'payableDate', 'dailyProfit', 'annualizedYield', 'totalUsdPlus', 'totalUsdc',
                 'duration']

    response = requests.get(url)
    transactionHash = [line['transactionHash'] for line in response.json()]
    payableDate = [line['payableDate'] for line in response.json()]
    dailyProfit = [line['dailyProfit'] for line in response.json()]
    annualizedYield = [line['annualizedYield'] for line in response.json()]
    totalUsdPlus = [line['totalUsdPlus'] for line in response.json()]
    totalUsdc = [line['totalUsdc'] for line in response.json()]
    duration = [line['duration'] for line in response.json()]

    # format --align--
    align = workbook.add_format()
    align.set_align('center')
    align.set_align('vcenter')

    # column width
    worksheet.set_column(0, 0, 75)
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(2, 2, 10)
    worksheet.set_column(3, 5, 15)
    worksheet.set_column(6, 6, 10)

    row = 0
    col = 0

    for line in colTitles:
        worksheet.write(row, col, line, align)
        col += 1
    row += 1

    for line in response.json():
        col = 0
        for value in line.values():
            worksheet.write(row, col, value)
            col += 1
        row += 1

    workbook.close()


net = ['O', 'M', 'B', 'A']


def start(network):
    oUrl = 'https://op.overnight.fi/api/dapp/payouts'
    mUrl = 'https://app.overnight.fi/api/dapp/payouts'
    bUrl = 'https://bsc.overnight.fi/api/dapp/payouts'
    aUrl = 'https://avax.overnight.fi/api/dapp/payouts'
    if network == net[0]:
        return write_data(oUrl)
    elif network == net[1]:
        return write_data(mUrl)
    elif network == net[2]:
        return write_data(bUrl)
    elif network == net[3]:
        return write_data(aUrl)
    else:
        print('Error')


start('A')
