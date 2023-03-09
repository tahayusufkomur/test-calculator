import os

import pandas as pd

from src.base_personality_test import BasePersonalityTest


class VarisPerformance(BasePersonalityTest):

    name = "Varis Performance"
    type = 2
    description = "Varisler için özel performans testi"
    sort_columns = ['Kendi adınızı seçiniz.']
    scoring = {
        1: 1,
        2: 2,
        3: 3,
        4: 4,
        5: 5
    }

    categories_dict = {1: ['Kişisel Özellikler', [i for i in range(1, 14)]],
                       2: ['Stratejik Bakış', [i for i in range(14, 19)]],
                       3: ['Eylemsel Beceriler', [i for i in range(19, 32)]],
                       4: ['İletişim Becerisi', [i for i in range(32, 39)]],
                       5: ['Acil Durum Becerisi', [i for i in range(39, 44)]],
                       6: ['Kişisel Donanım', [i for i in range(44, 51)]],
                       }

    should_be_same = [
    ]

    def format_report(self):

        column_1 = 'Kişisel Özellikler'
        column_2 = 'Stratejik Bakış'
        column_3 = 'Eylemsel Beceriler'
        column_4 = 'İletişim Becerisi'
        column_5 = 'Acil Durum Becerisi'
        column_6 = 'Kişisel Donanım'

        df = pd.read_excel(self.output_path, engine='openpyxl')

        mean_1 = df[column_1].mean()
        mean_2 = df[column_2].mean()
        mean_3 = df[column_3].mean()
        mean_4 = df[column_4].mean()
        mean_5 = df[column_5].mean()
        mean_6 = df[column_6].mean()
        mean_t_1 = df[column_1+" T"].mean()
        mean_t_2 = df[column_2+" T"].mean()
        mean_t_3 = df[column_3+" T"].mean()
        mean_t_4 = df[column_4+" T"].mean()
        mean_t_5 = df[column_5+" T"].mean()
        mean_t_6 = df[column_6 + " T"].mean()

        df = df.style \
            .applymap(self.adjust_font) \
            .applymap(lambda x: self.mean_highlighter(x, mean_1), subset=pd.IndexSlice[:, [column_1]])\
            .applymap(lambda x: self.mean_highlighter(x, mean_2), subset=pd.IndexSlice[:, [column_2]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_3), subset=pd.IndexSlice[:, [column_3]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_4), subset=pd.IndexSlice[:, [column_4]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_5), subset=pd.IndexSlice[:, [column_5]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_6), subset=pd.IndexSlice[:, [column_6]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_1), subset=pd.IndexSlice[:, [column_1+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_2), subset=pd.IndexSlice[:, [column_2+" T"]])  \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_3), subset=pd.IndexSlice[:, [column_3+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_4), subset=pd.IndexSlice[:, [column_4+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_5), subset=pd.IndexSlice[:, [column_5+" T"]])\
            .applymap(lambda x: self.mean_highlighter(x, mean_t_6), subset=pd.IndexSlice[:, [column_6 + " T"]])

        df.to_excel(self.output_path, engine='openpyxl', index=False)
        total_len = len(df.data)

        indexes = self.get_indexes()
        for i, index in enumerate(indexes):
            self.add_barchart("Sheet1",
                              start_col=column_1,
                              end_col=column_6,
                              name_col="Değerlendirilen vârisin adı-soyadını yazınız.",
                              output_path=self.output_path,
                              min_row=index['min_row'],
                              max_row=index['max_row']+1,
                              title=f"yönetici: {index['yönetici']}",
                              total=total_len,
                              multiplier=i+2)

    def create_report(self):
        self.filter_df_by_password()
        self.map_answers()
        # remove questions from the result and add total column
        self.prepare_result()
        # sort
        # self.sort_df()
        # add t-score
        self.add_tscore()
        # add mean
        self.add_mean_row()
        # add standart deviation
        self.add_std_row()
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
        # create text report depending on answers
        if os.path.exists(self.text_file_path):
            self.create_text_report()

    def sort_df(self):
        self.test_df.sort_values(self.sort_columns, inplace=True)
        self.result_df.sort_values(self.sort_columns, inplace=True)

    # takes performance df, and returns min row max row of manager - position
    def get_indexes(self):
        df = pd.read_excel(self.output_path, engine='openpyxl')
        df.drop(df.tail(2).index, inplace=True)
        manager_grouped = df.groupby('Kendi adınızı seçiniz.')
        min_row = 1
        graph_indexes = []
        for manager_name, manager_grp in manager_grouped:
            graph = {}
            graph['yönetici'] = manager_name
            graph['min_row'] = min_row + 1
            # TODO: check if this line works with multiple manager
            graph['max_row'] = min_row + len(manager_grouped) + 1 if len(manager_grouped) > 1 else min_row + 1
            graph_indexes.append(graph)
            min_row = min_row + len(manager_grouped)
        return sorted(graph_indexes, key=lambda d: d['yönetici'])


if __name__ == '__main__':
    x = VarisPerformance("/Users/tkoemue/ws/projects/merve/five-factor-test-calculator/yanıtlar/Varis.xlsx", None)
    x.create_report()