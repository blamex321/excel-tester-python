import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

data = pd.read_excel("/Users/blamex321/programming/python/dad-project/SL Agile Maturity - Self Assessment (Copy).xls")

for leader_name in data['Scrum Leader Name'].unique():
    leader_data = data[data['Scrum Leader Name'] == leader_name]
    with PdfPages(f'{leader_name}(self-assesment).pdf') as pdf_pages:
        yes_or_no_rows = [10, 11, 12, 13, 14, 15, 16, 19, 20, 21]
        question_rows = list(range(22, 57))

        def plot_number():
            for index, row in leader_data.iterrows():
                team_name = row[7]
                dev_no = row[17]
                time_diff = row[18]
                x_data=["Number of DEVS","Timezone Difference"]
                y_data=[dev_no,time_diff]
                colors=["blue","green"]
                plt.bar(x_data, y_data, color=colors)
                plt.title(f"Team Name:{team_name}")
                pdf_pages.savefig()
                plt.clf()

        def plot_question_answers(row_number):
            always, mostly, sometimes, rarely, never = 0, 0, 0, 0, 0
            x_data = ["always", "mostly", "sometimes", "rarely", "never"]
            y_data = []
            for index, row in leader_data.iterrows():
                if row[row_number] == 'Always':
                    always += 1
                elif row[row_number] == 'Mostly':
                    mostly += 1
                elif row[row_number] == 'Sometimes':
                    sometimes += 1
                elif row[row_number] == 'Rarely':
                    rarely += 1
                else:
                    never += 1
            y_data = [always, mostly, sometimes, rarely, never]
            colors = ['green', 'blue', 'yellow', 'violet', 'red']
            plt.bar(x_data, y_data, color=colors)
            plt.title(f"{leader_name}\n{leader_data.columns[row_number]}")
            plt.yticks(range(6), range(1,7))
            pdf_pages.savefig()
            plt.clf()

        def plot_yes_or_no_questions(row_number):
            yes_count, no_count = 0, 0
            x_data = ['yes', 'no']
            y_data = []
            for index, row in leader_data.iterrows():
                if row[row_number] == 'Yes':
                    yes_count += 1
                else:
                    no_count += 1
            y_data = [yes_count, no_count]
            colors = ['green', 'red']
            plt.bar(x_data, y_data, color=colors)
            plt.title(f"{leader_name}\n{leader_data.columns[row_number]}")
            plt.yticks(range(4), range(4))
            pdf_pages.savefig()
            plt.clf()

        for i in yes_or_no_rows:
            plot_yes_or_no_questions(i)
        for j in question_rows:
            plot_question_answers(j)
        plot_number()
