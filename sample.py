import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

data = pd.read_excel("SL Team Feedback (Copy)(1-7) (1).xlsx")
for leader_name in data['My Scrum Leader Name is '].unique():
    leader_data = data[data['My Scrum Leader Name is '] == leader_name]
    with PdfPages(f'{leader_name}(feedback).pdf') as pdf_pages:
        for index, row in leader_data.iterrows():
            team_name = row[6]
            daily = row[7]
            backlog=row[8]
            sprint = row[9]
            value=row[10]
            review=row[11]
            comments=row[12]
            if comments==None:
                comments="No Comments"

            x_data=["Daily Stand Up","Backlog Refinement","Sprint Planning","Value Delivery","Review & Adapt"]
            y_data=[daily,backlog,sprint,value,review]
            plt.rcParams["figure.figsize"] = [7.50, 3.50]
            plt.rcParams["figure.autolayout"] = True
            plt.bar(x_data, y_data)
            plt.title(f"Team Name:{team_name}")
            pdf_pages.savefig()
            plt.clf()

            # Add comments to a new page
            fig, ax = plt.subplots()
            ax.axis('off')
            ax.text(0.05, 0.5, f"Comments for Team Name: {team_name}\n\n{comments}")
            pdf_pages.savefig()
            plt.clf()
