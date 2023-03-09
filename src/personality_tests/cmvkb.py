import os

import pandas as pd
from src.base_personality_test import BasePersonalityTest


class Cmvkb(BasePersonalityTest):

    kuruma_baglilik = 27
    ic_motivasyon = 27

    name = "ÇMVKB"
    type = 2
    description = "Çalışan motivasyonu ve kurumsal bağlılık testi"

    scoring = {
        'Kesinlikle katılmıyorum': 1,
        'Katılmıyorum': 2,
        'Kararsızım': 3,
        'Katılıyorum': 4,
        'Kesinlikle katılıyorum': 5
    }

    text_scoring = {
        1: '1',
        2: '2',
        3: '3',
        4: '4',
        5: '5'
    }

    categories_dict = {1: ["Kuruma Bağlılık", [2, 4, 5, 6, 8, 10, 12, 13, 17, 19, 21]],
                       2: ['İç Motivasyon', [1, 3, 7, 9, 11, 14, 15, 16, 18, 20, 22]]
                       }
    categories = ["Kuruma Bağlılık", 'İç Motivasyon']

    def add_characteristic(self):
        self.result_df['Karakter'] = self.test_df.apply(self.map_characteristic, axis=1)

    def format_report(self):

        df = pd.read_excel(self.output_path, engine='openpyxl')
        loyality_mean = df['Kuruma Bağlılık'].mean()
        motivation_mean = df['İç Motivasyon'].mean()
        loyality_z_mean = df['Kuruma Bağlılık T'].mean()
        motivation_z_mean = df['İç Motivasyon T'].mean()

        df = df.style \
            .applymap(self.adjust_font) \
            .applymap(lambda x: self.mean_highlighter(x, loyality_mean), subset=pd.IndexSlice[:, ['Kuruma Bağlılık']])\
            .applymap(lambda x: self.mean_highlighter(x, motivation_mean), subset=pd.IndexSlice[:, ['İç Motivasyon']]) \
            .applymap(lambda x: self.mean_highlighter(x, loyality_z_mean), subset=pd.IndexSlice[:, ['Kuruma Bağlılık T']]) \
            .applymap(lambda x: self.mean_highlighter(x, motivation_z_mean), subset=pd.IndexSlice[:, ['İç Motivasyon T']]) \
            .applymap(lambda x: self.character_highlighter(x), subset=pd.IndexSlice[:, ['Karakter']]) \

        df.to_excel(self.output_path, engine='openpyxl', index=False)

        self.add_barchart("Sheet1", "Kuruma Bağlılık T", "İç Motivasyon T", "İsim Soyisim", self.output_path, multiplier=1)

    def character_highlighter(self, x):
        if x == 'Değer Katan, Çalışkan mutlular':
            return f"background-color: {self.colors['very_good']}; color: black;"
        elif x == 'Değer katmayan Yatan Mutlular':
            return f"background-color: {self.colors['bad']}; color: black;"
        elif x == 'Sorgulayan, Yenilikçiler':
            return f"background-color: {self.colors['good']}; color: black;"
        elif x == 'Gelenekseller, Olanı sürdürenler':
            return f"background-color: {self.colors['very_bad']}; color: black;"
        else:
            return None

    def prepare_result(self):
        self.result_df = self.test_df.copy()[self.personal_questions]
        for key, row in self.categories_dict.items():
            self.result_df[row[0]] = self.sum_row(row[1])
            self.test_df[row[0]] = self.sum_row(row[1])
        # will be used for getting highest score categories
        self.text_calculation_df = self.test_df.copy()

        self.kuruma_baglilik = self.result_df['Kuruma Bağlılık'].mean()
        self.ic_motivasyon = self.result_df['Kuruma Bağlılık'].mean()

        # add chacteristic
        self.add_characteristic()

    def map_characteristic(self, row):
        if row['Kuruma Bağlılık'] > self.kuruma_baglilik:
            if row['İç Motivasyon'] > self.ic_motivasyon:
                return 'Değer Katan, Çalışkan mutlular'
            else:
                return 'Değer katmayan Yatan Mutlular'
        else:
            if row['İç Motivasyon'] > self.ic_motivasyon:
                return 'Sorgulayan, Yenilikçiler'
            else:
                return 'Gelenekseller, Olanı sürdürenler'


if __name__ == '__main__':
    cm = Cmvkb('/Users/tkoemue/ws/projects/merve/five-factor-test-calculator/yanıtlar/ÇMVKB.xlsx', None)
    cm.create_report()
