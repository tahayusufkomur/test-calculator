import pandas as pd

from src.base_personality_test import BasePersonalityTest


class Cipto(BasePersonalityTest):

    name = "cipto"
    type = 2
    description = "Çalışanların iş yerindeki problemlere verdikleri tepkiler ölçeği"
    scoring = {
        'Bana hiç uygun değil': 1,
        'Bana uygun değil': 2,
        'Biraz uygun': 3,
        'Kısmen uygun': 4,
        'Oldukça uygun': 5,
        'Büyük ölçüde uygun': 6,
        'Tamamen uygun': 7
    }

    text_scoring = {
        1: '1',
        2: '2',
        3: '3',
        4: '4',
        5: '5',
        6: '6',
        7: '7'
    }

    reverse_scoring = {
        'Bana hiç uygun değil': 7,
        'Bana uygun değil': 6,
        'Biraz uygun': 5,
        'Kısmen uygun': 4,
        'Oldukça uygun': 3,
        'Büyük ölçüde uygun': 2,
        'Tamamen uygun': 1
    }

    categories_dict = {1: ['İstifa İhtimali', [1, 2, 3, 4, 5, 6]],
                       2: ['Düşünceli Konuşma', [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]],
                       3: ['Sabırlı Olma', [18, 19, 20, 21, 22]],
                       4: ['Saldırgan Konuşma', [23, 24, 25, 26, 27, 28, 29]],
                       5: ['İhmalkârlık', [30, 31, 32, 33, 34]],
                       }

    should_be_same = [
    ]

    # reverse_calculation_list = [
    #     1, 2, 3, 4, 5, 6,
    #     23, 24, 25, 26, 27, 28, 29,
    #     30, 31, 32, 33, 34
    # ]

    def format_report(self):

        column_1 = 'İstifa İhtimali'
        column_2 = 'Düşünceli Konuşma'
        column_3 = 'Sabırlı Olma'
        column_4 = 'Saldırgan Konuşma'
        column_5 = 'İhmalkârlık'

        df = pd.read_excel(self.output_path, engine='openpyxl')

        mean_1 = df[column_1].mean()
        mean_2 = df[column_2].mean()
        mean_3 = df[column_3].mean()
        mean_4 = df[column_4].mean()
        mean_5 = df[column_5].mean()
        mean_t_1 = df[column_1+" T"].mean()
        mean_t_2 = df[column_2+" T"].mean()
        mean_t_3 = df[column_3+" T"].mean()
        mean_t_4 = df[column_4+" T"].mean()
        mean_t_5 = df[column_5+" T"].mean()

        df = df.style \
            .applymap(self.adjust_font) \
            .applymap(lambda x: self.mean_highlighter(x, mean_1), subset=pd.IndexSlice[:, [column_1]])\
            .applymap(lambda x: self.mean_highlighter(x, mean_2), subset=pd.IndexSlice[:, [column_2]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_3), subset=pd.IndexSlice[:, [column_3]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_4), subset=pd.IndexSlice[:, [column_4]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_5), subset=pd.IndexSlice[:, [column_5]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_1), subset=pd.IndexSlice[:, [column_1+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_2), subset=pd.IndexSlice[:, [column_2+" T"]])  \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_3), subset=pd.IndexSlice[:, [column_3+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_4), subset=pd.IndexSlice[:, [column_4+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_5), subset=pd.IndexSlice[:, [column_5+" T"]])

        df.to_excel(self.output_path, engine='openpyxl', index=False)

        self.add_barchart("Sheet1", column_1, column_5, "İsim Soyisim", self.output_path)


if __name__ == '__main__':
    x = Cipto("/Users/tkoemue/ws/projects/merve/five-factor-test-calculator/yanıtlar/ÇİPTÖ.xlsx", None)
    x.create_report()