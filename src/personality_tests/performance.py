import os

import pandas as pd

from src.base_personality_test import BasePersonalityTest
from src.utilities import get_number_from_question


class Performance(BasePersonalityTest):

    name = "performance"
    type = 2
    description = "Çalışan performans değerlendirmesi"

    scoring = {
        'Hiç uygun değil': 1,
        'Uygun değil': 2,
        'Kararsızım': 3,
        'Uygun': 4,
        'Çok Uygun': 5
    }

    reverse_scoring = {
        'Hiç uygun değil': 5,
        'Uygun değil': 4,
        'Kararsızım': 3,
        'Uygun': 2,
        'Çok Uygun': 1
    }

    text_scoring = {
        'Hiç uygun değil': '1',
        'Uygun değil': '2',
        'Kararsızım': '3',
        'Uygun': '4',
        'Çok Uygun': '5'
    }

    reverse_calculation_list = [
        5, 14, 19, 20, 21, 25, 29, 32, 34, 40, 51, 52, 53
    ]

    categories_dict = {1: ["Teknik iş bilgisi ve donanımı", [i for i in range(1, 4)]],
                       2: ['İnisiyatif alabilme ve muhakeme yeteneği', [i for i in range(4, 9)]],
                       3: ['İşbirliği ve takımdaşlık', [i for i in range(9, 15)]],
                       4: ['İş yapış şekli ve sorumluluk bilinci', [i for i in range(15, 29)]],
                       5: ['Eğitim ve gelişime açıklık', [i for i in range(29, 37)]],
                       6: ['Süreç Yönetimi ve Kalite & Kontrol becerisi', [i for i in range(37, 45)]],
                       7: ['İş güvenliği', [i for i in range(45, 47)]],
                       8: ['Kurum Kültürüne Uyum ve İlişkilere Hassasiyet', [i for i in range(47, 55)]]
                       }

    def map_answers(self):
        for key, row in self.categories_dict.items():
            for column in self.test_questions:
                if get_number_from_question(column) in row[1]:
                    if get_number_from_question(column) in self.reverse_calculation_list:
                        self.test_df[column] = self.test_df[column].map(self.reverse_scoring)
                    else:
                        self.test_df[column] = self.test_df[column].map(self.scoring)

    def create_report(self):
        # create text report depending on answers
        if os.path.exists(self.text_file_path):
            self.create_text_report()
        # map answers to their corresponding score in dataframe
        self.map_answers()
        # remove questions from the result and add total column
        self.prepare_result()
        # sort
        self.sort_df()
        # add mean
        self.add_mean_row()
        # add standart deviation
        self.add_std_row()
        # add z-score
        self.add_tscore()
        # round df
        self.round_df()
        # save
        self.save_report()

        # create 1 copy to creating colors and adjusting width
        self.copy_df = self.result_df.copy()
        # colorize
        self.format_report()
        # adjust width
        self.adjust_width()
        print(f"Created report for: {self.name}")

    def format_report(self):

        df = pd.read_excel(self.output_path, engine='openpyxl')
        category_1 = df['Teknik iş bilgisi ve donanımı'].mean()
        category_2 = df['İnisiyatif alabilme ve muhakeme yeteneği'].mean()
        category_3 = df['İşbirliği ve takımdaşlık'].mean()
        category_4 = df['İş yapış şekli ve sorumluluk bilinci'].mean()
        category_5 = df['Eğitim ve gelişime açıklık'].mean()
        category_6 = df['Süreç Yönetimi ve Kalite & Kontrol becerisi'].mean()
        category_7 = df['İş güvenliği'].mean()
        category_8 = df['Kurum Kültürüne Uyum ve İlişkilere Hassasiyet'].mean()
        category_1_t = df['Teknik iş bilgisi ve donanımı T'].mean()
        category_2_t = df['İnisiyatif alabilme ve muhakeme yeteneği T'].mean()
        category_3_t = df['İşbirliği ve takımdaşlık T'].mean()
        category_4_t = df['İş yapış şekli ve sorumluluk bilinci T'].mean()
        category_5_t = df['Eğitim ve gelişime açıklık T'].mean()
        category_6_t = df['Süreç Yönetimi ve Kalite & Kontrol becerisi T'].mean()
        category_7_t = df['İş güvenliği T'].mean()
        category_8_t = df['Kurum Kültürüne Uyum ve İlişkilere Hassasiyet T'].mean()

        df = df.style \
            .applymap(self.adjust_font) \
            .applymap(lambda x: self.mean_highlighter(x, category_1), subset=pd.IndexSlice[:, ['Teknik iş bilgisi ve donanımı']])\
            .applymap(lambda x: self.mean_highlighter(x, category_2), subset=pd.IndexSlice[:, ['İnisiyatif alabilme ve muhakeme yeteneği']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_3), subset=pd.IndexSlice[:, ['İşbirliği ve takımdaşlık']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_4), subset=pd.IndexSlice[:, ['İş yapış şekli ve sorumluluk bilinci']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_5), subset=pd.IndexSlice[:, ['Eğitim ve gelişime açıklık']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_6), subset=pd.IndexSlice[:, ['Süreç Yönetimi ve Kalite & Kontrol becerisi']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_7), subset=pd.IndexSlice[:, ['İş güvenliği']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_8), subset=pd.IndexSlice[:, ['Kurum Kültürüne Uyum ve İlişkilere Hassasiyet']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_1_t), subset=pd.IndexSlice[:, ['Teknik iş bilgisi ve donanımı T']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_2_t), subset=pd.IndexSlice[:, ['İnisiyatif alabilme ve muhakeme yeteneği T']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_3_t), subset=pd.IndexSlice[:, ['İşbirliği ve takımdaşlık T']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_4_t), subset=pd.IndexSlice[:, ['İş yapış şekli ve sorumluluk bilinci T']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_5_t), subset=pd.IndexSlice[:, ['Eğitim ve gelişime açıklık T']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_6_t), subset=pd.IndexSlice[:, ['Süreç Yönetimi ve Kalite & Kontrol becerisi T']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_7_t), subset=pd.IndexSlice[:, ['İş güvenliği T']]) \
            .applymap(lambda x: self.mean_highlighter(x, category_8_t), subset=pd.IndexSlice[:, ['Kurum Kültürüne Uyum ve İlişkilere Hassasiyet T']]) \

        df.to_excel(self.output_path, engine='openpyxl', index=False)
        indexes = self.get_indexes()

        total_len = len(df.data)

        self.add_barchart("Sheet1",
                          start_col="Teknik iş bilgisi ve donanımı T",
                          end_col="Kurum Kültürüne Uyum ve İlişkilere Hassasiyet T",
                          name_col="Değerlendirmesini yaptığınız personeli seçiniz.",
                          output_path=self.output_path,
                          min_row=2,
                          title="HEPSİ")
        # add barcharts
        for i, index in enumerate(indexes):
            self.add_barchart("Sheet1",
                              start_col="Teknik iş bilgisi ve donanımı T",
                              end_col="Kurum Kültürüne Uyum ve İlişkilere Hassasiyet T",
                              name_col="Değerlendirmesini yaptığınız personeli seçiniz.",
                              output_path=self.output_path,
                              min_row=index['min_row'],
                              max_row=index['max_row'],
                              title=f"yönetici: {index['yönetici']} - değerlendirilen pozisyon: {index['pozisyon']}",
                              total=total_len,
                              multiplier=i+2)

    def sort_df(self):
        self.test_df.sort_values(['Değerlendiren kişi: ', 'Değerlendirdiğiniz kişinin ünvanını seçiniz.  '], inplace=True)
        self.result_df.sort_values(['Değerlendiren kişi: ', 'Değerlendirdiğiniz kişinin ünvanını seçiniz.  '], inplace=True)

    def prepare_result(self):
        self.result_df = self.test_df.copy()[self.personal_questions]
        for key, row in self.categories_dict.items():
            self.result_df[row[0]] = (self.sum_row(row[1])/len(row[1]))*20
            self.test_df[row[0]] = (self.sum_row(row[1])/len(row[1]))*20
        # will be used for getting highest score categories
        self.text_calculation_df = self.test_df.copy()

# takes performance df, and returns min row max row of manager - position
    def get_indexes(self):
        df = pd.read_excel(self.output_path, engine='openpyxl')
        df.drop(df.tail(2).index, inplace=True)
        manager_grouped = df.groupby('Değerlendiren kişi: ')
        min_row = 1
        graph_indexes = []
        for manager_name, manager_grp in manager_grouped:
            position_grouped = manager_grp.groupby('Değerlendirdiğiniz kişinin ünvanını seçiniz.  ')
            for position_name, position_grp in position_grouped:
                graph = {}
                graph['yönetici'] = manager_name
                graph['pozisyon'] = position_name
                graph['min_row'] = min_row + 1
                graph['max_row'] = min_row + len(position_grp) + 1
                graph_indexes.append(graph)
                min_row = min_row + len(position_grp)
        return sorted(graph_indexes, key=lambda d: d['pozisyon'])


if __name__ == '__main__':
    cm = Performance('/Users/tkoemue/ws/projects/merve/five-factor-test-calculator/yanıtlar/Performance.xlsx', None)
    cm.create_report()
