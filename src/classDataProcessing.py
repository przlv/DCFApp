import pandas as pd
import numpy as np
import numpy_financial as npf


class DataProccessing:
    
    def __init__(self,  filename, with_year, by_year) -> None:

        self.dataExcel = pd.ExcelFile(filename)
        self.annual_interval = [with_year, by_year] if type(with_year) == int else [int(with_year), int(by_year)]


    def run(self) -> pd.DataFrame:
        
        self._indicators_generation()
        print(self.indicators)


    def _indicators_generation(self):
        dict_years_example = generate_dictionary(self.annual_interval[0], self.annual_interval[1])
        
        dict_cashFlow = {
            "Инвестиции по проекту": dict_years_example,
            "Совокупное изменение в доходах и расходах": dict_years_example,
            "Амортизация": dict_years_example,
            "Налог на прибыль": dict_years_example,
            "Налог на имущество": dict_years_example,
            "Денежный поток": dict_years_example,
            "Накопленный денежный поток": dict_years_example,
            "Фактор дисконтирования": dict_years_example,
            "Дисконтированный денежный поток": dict_years_example,
            "Накопленный дисконтированный денежный поток": dict_years_example,
            "Вспомогательно: расчет DPP": dict_years_example,
            "Вспомогательно: расчет PP": dict_years_example,
        }

        # Основные показатели проекта:		
        current_list = self.dataExcel.parse('Денежн. поток на буд. ст-ть')
       
        count = 0
        for row in current_list.itertuples():
            if count == 11: break
            row = list(row)
            if row[2] in list(dict_cashFlow.keys()):
                dict_cashFlow[row[2]] = row
                count+=1                
        
        self.indicators = {
            'NPV': None,
            'DPI': None,
            'IRR': None,
            'PP': None,
            'DPP': None,
        }
        
        indic_sum = 0
        for val in dict_cashFlow['Дисконтированный денежный поток'][4:-1]:
            indic_sum += val
        self.indicators['NPV'] = indic_sum
        
        indic_sum = 0
        from math import isnan
        for val1, val2 in zip(dict_cashFlow['Инвестиции по проекту'][4:-1], dict_cashFlow['Фактор дисконтирования'][4:-1]):
            if not isnan(val1) and not isnan(val2):
                indic_sum += val1* val2
        self.indicators['DPI'] = self.indicators['NPV']/(indic_sum) + 1
        
        if dict_cashFlow['Денежный поток'][4] >= 0:
            self.indicators['IRR'] = 1
        else:
            rate = npf.irr(dict_cashFlow['Денежный поток'][4:-1])
            self.indicators['IRR'] = rate
        
        indic_sum = 0
        for val in dict_cashFlow['Вспомогательно: расчет DPP'][4:-1]:
            indic_sum += val
        if indic_sum >= 20:
            self.indicators['PP'] = 'Не окупается'
        else:
            self.indicators['PP'] = indic_sum
        
        indic_sum = 0
        for val in dict_cashFlow['Вспомогательно: расчет PP'][4:-1]:
            indic_sum += val
        if indic_sum >= 20:
            self.indicators['DPP'] = 'Не окупается'
        else:
            self.indicators['DPP'] = indic_sum
        
        return


def generate_dictionary(start_year, end_year):
    dictionary = {}
    for year in range(start_year, end_year+1):
        dictionary[year] = None
    return dictionary


filename = 'data/data.xlsx'
start_year = '2023'
end_year = '2041'

if __name__ == '__main__':
    testapp = DataProccessing(filename, start_year, end_year)
    
    testapp.run()
    