import requests
from openpyxl import load_workbook
from exel_list import write_to_sheet_1, sheet

arr = []
def find_item(response, id_articul: int):
    date = response.json()
    products = date['products']
    for index, value in enumerate(products):
        if int(value['id']) == id_articul:
            return index + 1

    return None

def get_resp():

    # Открываем файл
    #wb = load_workbook("shorts_position_words.xlsx")

    # Выбираем активный лист
    #ws = wb.active
    exel = sheet.col_values(1)
    # Считываем данные из столбца J (10-й столбец)
    #keywords = [cell.value for cell in ws["A"] if cell.value is not None]
    keywords = [cell for cell in exel if cell]

    for keyword in keywords:
        page = 1
        while True:
            cookies = {
                'external-locale': 'ru',
                '_wbauid': '8264840631760617509',
                'wbx-validation-key': '0b839a54-4dfe-4b2a-a0a6-edf2ba0b6db4',
                'x-supplier-id-external': '0cd68178-a8ec-409f-8817-787cd8da5b16',
                '_cp': '1',
                'x_wbaas_token': '1.1000.2de0937020904584ad4130f6ea4d9513.MHwxODguMTIwLjk5LjE3NXxNb3ppbGxhLzUuMCAoV2luZG93cyBOVCAxMC4wOyBXaW42NDsgeDY0KSBBcHBsZVdlYktpdC81MzcuMzYgKEtIVE1MLCBsaWtlIEdlY2tvKSBDaHJvbWUvMTM4LjAuMC4wIFlhQnJvd3Nlci8yNS44LjAuMCBTYWZhcmkvNTM3LjM2fDE3NjQwMTQ0MjN8cmV1c2FibGV8MnxleUpvWVhOb0lqb2lJbjA9fDB8M3wxNzYzNDA5NjIz.MEUCIQCgUE34SkyWarO2CKg1aQO2RhTrfO0BYeQrB9s8HtNHGQIgfiSVrBltmTmFUJGneep0oWfHTGiMAh5YsmVX/QCLsjo=',
                '__zzatw-wb': 'MDA0dC0cTHtmcDhhDHEWTT17CT4VHThHKHIzd2UqP2YlaE5eJjVRP0FaW1Q4NmdBEXUmCQg3LGBwVxlRExpceEdXeiwfE390I08QEmQ/SGllbQwtUlFRS19/Dg4/aU5ZQ11wS3E6EmBWGB5CWgtMeFtLKRZHGzJhXkZpdRVSCxFeQnN0dCw9bFBjShVQdBAICFpNRXkmKk8NFF1BR19vG3siXyoIJGM1Xz9EaVhTMCpYQXt1J3Z+KmUzPGwjYU9iH0NeVAooGg1pN2wXPHVlLwkxLGJ5MVIvE0tsP0caRFpbQDsyVghDQE1HFF9BWncyUlFRS2EQR0lrZU5TQixmG3EVTQgNND1aciIPWzklWAgSPwsmIBd5cyxPfxRiRkRxbxt/Nl0cOWMRCxl+OmNdRkc3FSR7dSYKCTU3YnAvTCB7SykWRxsyYV5GaXUVTzo/YUVCdHsmbG1SGkRdH0wTSgktGhh0citWOj9jcXJyLSpBVxlRDxZhDhYYRRcje0I3Yhk4QhgvPV8/YngiD2lIYCVFXVZ5JSIZem8kS3FPLH12X30beylOIA0lVBMhP05yQUspmw==',
                'cfidsw-wb': 'OEm0RZfcZEsg83CUc7njOuPeB8DYr6ds2ppwpNqAdo2wujXPsTuc8BWRbRsEgJMaTSkjWnIP4GPzMYR/eC+5Pqg93UJYfAMFoiJXyjMG6xob07gEHEiKGRtigB+UngdlvKFJpQb8U/5fCTf3YmNqWXSougHWlna31HWtIwk=',
            }

            params = {
                'ab_testing': [
                    'false',
                    'false',
                ],
                'appType': '1',
                'curr': 'rub',
                'dest': '-1257786',
                'hide_dtype': '11',
                'inheritFilters': 'false',
                'lang': 'ru',
                'page': str(page),
                'query': str(keyword),
                'resultset': 'catalog',
                'sort': 'popular',
                'spp': '30',
                'suppressSpellcheck': 'false',
            }

            response = requests.get(
                'https://www.wildberries.ru/__internal/u-search/exactmatch/ru/common/v18/search',
                params=params,
                cookies=cookies,
                #headers=headers,
            )

            data = response.json()
            if not data.get('products'):
                break  # если товаров больше нет — выходим из while


            #if len(response.json()) != 0:
            result = find_item(response, id_articul=259711529) #259711529 78035685
            if result:
                arr.append(str((page - 1) * 100 + result))
                break
            page += 1
    #запись в exel
    write_to_sheet_1(arr)

def main():
    get_resp()


if __name__ == '__main__':
    main()

