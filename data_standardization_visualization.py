# Data Standardization and Visualization


# Prepare the Output Directories

META_ANALYSIS_PARENT_PATH = r"C:\Users\ma\Documents\Productive & Career\Internships\RA Kevin Boudreau\Research\Platform Meta Analysis"

def make_output_directory(meta_analysis_parent_path, output_directory):
    os.chdir(f'{meta_analysis_parent_path}')
    os.makedirs(output_directory, exist_ok=False)
    
    print(f'Output Directory: {output_directory}')
    

def make_intensive_extensive_directory(meta_analysis_parent_path, output_directory, intensive):
    os.chdir(f'{meta_analysis_parent_path}/{output_directory}')
    
    directory_name = "Intensive" if intensive else "Extensive"    

    os.makedirs(directory_name, exist_ok=False)
    os.chdir(directory_name)
    
    print(f'Currently Working in: {directory_name}')
    

# Extract Data From Excel

EXCEL_SHEET_PATH = r"All Tallies.xlsx"

PLATFORM_NAMES = ["Open Source", "App Developers", "Wikipedia", "Crowdsourcing", "Citizen Science"]

import pandas as pd

def parse_all_tallies(file_path, meta_analysis_parent_path, names, intensive):
    file = pd.ExcelFile(f'{meta_analysis_parent_path}/{file_path}')
    parsed_sheets = []
    total_number_of_studies = []
    for name in names:
        sheet = file.parse(f'Data - {name} ({"In." if intensive else "Ex."})')
        parsed_sheets.append([name, sheet])
        total_number_of_studies.append(sheet.shape[0])
    
    return parsed_sheets, total_number_of_studies


# Move to CSV

import csv
import re
import math
import os

MOTIVATION_COLUMNS = [[8, "Monetary"], [9, "Own-Use"], [10, "Learning"], [13, "Task-Related"], [14, "Autonomy"], [15, "Self-Efficacy"], [16, "Ideology"], [17, "Identity"], [20, "Peer-Status"], [21, "Career-Concerns"], [22, "Community"], [23, "Reciprocity"], [24, "Altruism"]]

CSV_PATH_NAMES = ["TemporaryOpenSource.csv", "TemporaryAppDevelopers.csv", "TemporaryWikipedia.csv", "TemporaryCrowdsourcing.csv", "TemporaryCitizenScience.csv"]


def convert_to_csv(parsed_platform, motivation, csv_path, intensive):
    platform_name = parsed_platform[0]
    platform_content = parsed_platform[1]
    
    motivation_column = motivation[0]
    motivation_name = motivation[1]
    
    num_rows = platform_content.shape[0]
    
    kept_rows = [["study", "mean", "sd", "n"]]
    
    os.makedirs("Temporary CSV Files", exist_ok=True)
    
    for r in range(num_rows):
        
        box = platform_content.iloc[r][motivation_column]
        
        # For a first step of standardization, convert all data to a 0 to 1 scale
        if not pd.isna(box):
            data_values = re.findall('(\d+\.\d+|\d+)', str(box))
            mean_value = float(data_values[0])
            sd_value = (float(data_values[1]) + float(data_values[2])) / 2 if len(data_values) == 3 else float(data_values[1]) if len(data_values) >= 2 else None
            
            study_name = platform_content.iloc[r][3]
            study_names_list = study_name.split(", ")
            study_name_short = study_names_list[-1]
            
            data_type = str(platform_content.iloc[r][5])
            
            sample_numbers = re.findall('\d+', str(platform_content.iloc[r][6]))
            sample_size = float(sample_numbers[0]) if sample_numbers else None
            
            if (intensive):
                
                if ("Likert (1-5)" in data_type):
                    mean_standardized = (mean_value - 1) / 4
                    sd_standardized = sd_value / 4
                    
                elif ("Likert (1-7)" in data_type):
                    mean_standardized = (mean_value - 1) / 6
                    sd_standardized = sd_value / 6
                    
                elif ("Likert (5-1)" in data_type):
                    mean_standardized = 1 - ((5 - mean_value) / 4)
                    sd_standardized = sd_value / 4
                    
                elif ("Scale (1-5)" in data_type):
                    mean_standardized = (mean_value - 1) / 4
                    sd_standardized = sd_value / 4
                    
                elif ("Scale (1-6)" in data_type):
                    mean_standardized = (mean_value - 1) / 5
                    sd_standardized = sd_value / 5
                    
                elif ("Scale (1-7)" in data_type):
                    mean_standardized = (mean_value - 1) / 6
                    sd_standardized = sd_value / 6
                    
                elif ("Number of Respondents" in data_type):
                    raise Exception("Number of Respondents should not be included in the analysis")
                    
                elif ("Percentage of Respondents" in data_type):
                    if ("Strongly-Agree" in str(box) or "Agree" in str(box) or "Unsure" in str(box) or "Disagree" in str(box) or "Strongly-Disagree" in str(box)):
                        mean_value = (
                            5 * float(data_values[0]) / 100 * sample_size +
                            4 * float(data_values[1]) / 100 * sample_size +
                            3 * float(data_values[2]) / 100 * sample_size +
                            2 * float(data_values[3]) / 100 * sample_size +
                            1 * float(data_values[4]) / 100 * sample_size
                        ) / sample_size
                        
                        sd_value = math.sqrt(
                            (
                                5 ** 2 * float(data_values[0]) / 100 * sample_size +
                                4 ** 2 * float(data_values[1]) / 100 * sample_size +
                                3 ** 2 * float(data_values[2]) / 100 * sample_size +
                                2 ** 2 * float(data_values[3]) / 100 * sample_size +
                                1 ** 2 * float(data_values[4]) / 100 * sample_size
                            ) / sample_size
                            - mean_value ** 2
                        )
                        
                        mean_standardized = (mean_value - 1) / 4
                        sd_standardized = sd_value / 4
                        
                    else:
                        raise Exception("Binomial Percentage of Respondents should not be included in the analysis")
                    
                else:
                    raise Exception(f"{data_type} should not be included in the analysis")
               
            #############################################################################################################################
                
            else:
                # Mention in paper that standard deviation was calculated using the formula for a binomial (though it's not always binomial)
                if ("Number of Respondents" in data_type):
                    mean_standardized = mean_value / sample_size
                    sd_standardized = math.sqrt(mean_standardized * (1 - mean_standardized))
                    
                elif ("Percentage of Respondents" in data_type):
                    
                    if ("Strongly-Agree" in str(box) or "Agree" in str(box) or "Unsure" in str(box) or "Disagree" in str(box) or "Strongly-Disagree" in str(box)):
                        raise Exception("Multinomial Percentage of Respondents should not be included in the analysis")
                    
                    # We can't be certain that all who don't say Work are necessarily Non-Work, so the max is the best bet
                    elif ("Non-Work" in str(box) or "Work" in str(box)):
                        mean_value = max(float(data_values[0]), float(data_values[1]))
                    
                    mean_standardized = mean_value / 100
                    sd_standardized = math.sqrt(mean_standardized * (1 - mean_standardized))
                    
                else:
                    raise Exception(f"{data_type} should not be included in the analysis")
            
            print(f'{platform_name} - {motivation_name}: {study_name_short} Added {[study_name_short, mean_standardized, sd_standardized, sample_size]}')
            kept_rows.append([study_name_short, mean_standardized, sd_standardized, sample_size])
    
    with open(f'Temporary CSV Files/{csv_path}', 'w', newline='') as new_platform:
        new_platform_writer = csv.writer(new_platform)
        new_platform_writer.writerows(kept_rows)


# Meta Analysis for Individual Motivations for Individual Platforms

import subprocess

R_SCRIPT_EXECUTABLE_PATH = r"C:\Program Files\R\R-4.3.1\bin\Rscript"
R_FILE_PATH = r"C:\Users\ma\Documents\Productive & Career\Internships\RA Kevin Boudreau\Research\Platform Meta Analysis\MetaAnalysis.R"

def generate_platform_motivation_summaries(platform, motivation_number, motivation, platform_path, r_script_path, r_file_path):
    os.makedirs(platform, exist_ok=True)
    
    r_code = \
        f'''
        # Load the Library
        library(metafor)

        # Read Data
        data <- read.csv("Temporary CSV Files/{platform_path}")

        # Variance
        data$var <- (data$sd^2) / (data$n)

        # DerSimonian and Laird
        res <- rma(yi = data$mean, vi = data$var, slab = data$study, method = "DL")

        # Statistics
        sink("{platform}/{platform}___{motivation_number}___{motivation}___Output.txt")
        print(res)
        total_sample_size <- sum(data$n)
        print(paste("Total Sample Size: ", total_sample_size))
        total_studies <- nrow(data)
        print(paste("Total Number of Studies: ", total_studies))
        sink()

        # Save the Forest Plot to a PNG file
        png("{platform}/{platform}___{motivation_number}___{motivation}.png")
        forest(res)
        dev.off()
        '''
    
    with open(r_file_path, "w") as r_file:
        r_file.write(r_code)
    
    result = subprocess.run([r_script_path, r_file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        print(f'Error occurred: {result.stderr.decode()}')
    else:
        print(f'Script output: {result.stdout.decode()}')
        
        new_platform_motivation_summary_path = f'{platform}/{platform}___{motivation_number}___{motivation}___Output.txt'
        
        print(new_platform_motivation_summary_path)
        
        return new_platform_motivation_summary_path


# Collect Summary Statistics for All Motivations

PLATFORM_SUMMARY_PATHS = [r"OpenSource.csv", r"AppDevelopers.csv", r"Wikipedia.csv", r"Crowdsourcing.csv", r"CitizenScience.csv"]
MOTIVATION_SUMMARY_PATHS = [r"Monetary.csv", r"Own-Use.csv", r"Learning.csv", r"Task-Related.csv", r"Autonomy.csv", r"Self-Efficacy.csv", r"Ideology.csv", r"Identity.csv", r"Peer-Status.csv", r"Career-Concerns.csv", r"Community.csv", r"Reciprocity.csv", r"Altruism.csv"]
CUMULATIVE_TABLE_PATH = "Cumulative Table.csv"


def generate_table_labels(cumulative_table, motivations, platforms):
    cumulative_table[0].extend(motivation[1] for motivation in motivations)
    cumulative_table.extend([platform] + ["" for space in range(len(cumulative_table[0]) - 1)] for platform in platforms)
    return cumulative_table


def save_cumulative_table(cumulative_table, cumulatve_table_path):
    with open(cumulatve_table_path, 'w') as cumulative_table_file:
        cumulative_table_file_writer = csv.writer(cumulative_table_file)
        cumulative_table_file_writer.writerows(cumulative_table)


def get_summary_statistics(motivation_number, platform_number, motivation, platform, platform_motivation_summary_path, platform_summary_path, motivation_summary_path, cumulative_table):
    os.makedirs("Platforms", exist_ok=True)
    os.makedirs("Motivations", exist_ok=True)
    
    try:
        with open(platform_motivation_summary_path, 'r') as platform_motivation_summary:
            lines = platform_motivation_summary.readlines()

        line_1 = lines[14]
        print(line_1)
        data_mean_se = re.findall(r'\d+\.\d+', line_1)[0:2]
        
        line_2 = lines[19]
        print(line_2)
        data_sample_size = re.findall(r'\d+', line_2)
        
        line_3 = lines[20]
        print(line_3)
        data_total_number_of_studies = re.findall(r'\d+', line_3)
        
        data = [data_mean_se[0], data_mean_se[1], data_sample_size[1], data_total_number_of_studies[1]]
        print(data)
        
        cumulative_table[platform_number + 1][motivation_number + 1] = f'{data[0]} ({data[1]}) {line_1.split()[-1]} (n = {data_sample_size[1]}) (N = {data_total_number_of_studies[1]})'
    
    except:
        
        cumulative_table[platform_number + 1][motivation_number + 1] = f'NOT MENTIONED'
        
        return

    platform_already_existed = os.path.isfile(f'Platforms/{platform_summary_path}')
    motivation_already_existed = os.path.isfile(f'Motivations/{motivation_summary_path}')
    
    with open(f'Platforms/{platform_summary_path}', 'a') as platform_summary:
        platform_summary_writer = csv.writer(platform_summary)
        
        if not platform_already_existed:
            platform_summary_writer.writerow(["Motivation", "mean", "se", "n", "studies"])
        
        platform_data = [motivation]
        
        platform_data.extend(data)
        
        print(platform_data)
        
        platform_summary_writer.writerow(platform_data)
    
    with open(f'Motivations/{motivation_summary_path}', 'a') as motivation_summary:
        motivation_summary_writer = csv.writer(motivation_summary)
        
        if not motivation_already_existed:
            motivation_summary_writer.writerow(["Platform", "mean", "se", "n", "studies"])
        
        motivation_data = [platform]
        
        motivation_data.extend(data)
        
        motivation_summary_writer.writerow(motivation_data)


# Meta Analysis for Individual Platforms and Individual Motivations

def generate_summary(summary_type, opposing_summary_type, summary, summary_path, r_script_path, r_file_path):
    os.makedirs(summary_type, exist_ok=True)
    
    print(f'{summary_type}/{summary_path}')
    
    r_code = \
        f'''
        # Load the Library
        library(metafor)

        # Read Data
        data <- read.csv("{summary_type}/{summary_path}")

        # Variance
        data$var <- (data$se^2)

        # DerSimonian and Laird
        res <- rma(yi = data$mean, vi = data$var, slab = paste(data${opposing_summary_type}, "( N =", data$studies, ")"), method = "DL")

        # Statistics
        sink("{summary_type}/{summary}___Output.txt")
        print(res)
        sink()

        # Save the Forest Plot to a PNG file
        png("{summary_type}/{summary}.png")
        forest(
            res,
            xlim = c(-0.75, 1.5),
            at = seq(0, 1, by = 0.1)
        )
        dev.off()
        '''
    
    with open(r_file_path, "w") as r_file:
        r_file.write(r_code)
    
    result = subprocess.run([r_script_path, r_file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        print(f'Error occurred: {result.stderr.decode()}')
    else:
        print(f'Script output: {result.stdout.decode()}')
        
        new_platform_motivation_summary_path = f'{summary_type}/{summary}___Output.txt'
        
        print(new_platform_motivation_summary_path)
        
        return new_platform_motivation_summary_path
        

def update_table_data(cumulative_table, row, column, summary_path):
    
    if (summary_path == None):
        print(f'Summary path does not exist as no studies mentioned the corresponding motivation')
        
        return
    
    with open(summary_path, 'r') as summary:
        lines = summary.readlines()
    
    significance = lines[14].split()[6] if len(lines[14].split()) == 7 else ''
        
    cumulative_table[row][column] += f' {significance}'
    

def record_total_number_of_studies(cumulative_table, platform_number, total_number_of_studies):
    cumulative_table[platform_number + 1][0] += f' (N = {total_number_of_studies[platform_number]})'


# Generate a Spider Graph and Column Chart

import matplotlib.pyplot as plt
import numpy as np

SPIDER_PATH = "Cumulative Spider.png"
COLUMN_PATHS = ["Extrinsic Column.png", "Intrinsic Column.png", "Social Column.png"]

def generate_plots(cumulative_table_csv_path, spider_path, column_paths, meta_analysis_parent_path, output_directory, intensive):
    
    directory_name = "Intensive" if intensive else "Extensive"
    
    os.chdir(f'{meta_analysis_parent_path}/{output_directory}/{directory_name}')
    
    intensive_label = '(Intensive)' if intensive else '(Extensive)'
    
    data = []
    with open(cumulative_table_csv_path, 'r') as cumulative_table_file:
        cumulative_table_file_reader = csv.reader(cumulative_table_file)
        for row in cumulative_table_file_reader:
            if row:
                data.append(row)
    
    categories = [category.split()[0] for category in data[0][1:]]

    values = [[float(box.split()[0]) if "NOT MENTIONED" not in box else 0 for box in data[row][1:]] for row in range(1, len(data))]
    
    standard_errors = [[float(box.split()[1][1:-1]) if "NOT MENTIONED" not in box else 0 for box in data[row][1:]] for row in range(1, len(data))]
    
    n_values = [[float(box.split()[5][:-1]) if "NOT MENTIONED" not in box else 0 for box in data[row][1:]] for row in range(1, len(data))]
    
    N_values = [[int(box.split()[8][:-1]) if "NOT MENTIONED" not in box else 0 for box in data[row][1:]] for row in range(1, len(data))]
    
    total_number_of_studies = [int(re.findall(r'\d+', data[row][0])[0]) for row in range(1, len(data))]
    
    #labels = [platform[0] for platform in data[1:]]
    labels = ["Open Source", "App Developers", "Wikipedia", "Crowdsourcing", "Citizen Science"]

    # Create a spider graph
    num_vars = len(categories)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()

    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))

    # Helper function to plot each individual
    def add_to_spider(values, label, ax, color):
        # Only include angles/values where values are not None
        valid_values = [(angle, value) for angle, value in zip(angles, values)]
        
        valid_angles = [item[0] for item in valid_values]
        valid_values = [item[1] for item in valid_values]

        # Close the shape by appending the first valid value and angle
        valid_angles.append(valid_angles[0])
        valid_values.append(valid_values[0])
        
        ax.plot(valid_angles, valid_values, color=color, linewidth=2, label=label)
        ax.fill(valid_angles, valid_values, color=color, alpha=0.25)


    # Plot each row of the dataset
    colors = ['blue', 'green', 'red', 'purple', 'orange']
    for val, label, color in zip(values, labels, colors):
        add_to_spider(val, label, ax, color)

    # Set the ytick labels to be empty
    ax.set_yticklabels([])
    ax.set_xticks(angles)
    ax.set_xticklabels(categories)

    plt.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1))
    plt.title(f"Spider Graph {intensive_label}")
    plt.savefig(spider_path, bbox_inches='tight')
    
    
    
    # Column Chart
    def generate_column_chart_with_error_bars(labels, values, standard_errors, categories, column_chart_path, title):
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Width of a bar 
        width = 0.15
        positions = np.arange(len(categories))
        
        for idx, (val, err, label, color) in enumerate(zip(values, standard_errors, labels, ['blue', 'green', 'red', 'purple', 'orange'])):
            ax.bar(positions + idx * width, val, width, label=label, color=color, yerr=err, capsize=5, align='center')

        ax.set_xlabel('Categories')
        ax.set_ylabel('Values')
        ax.set_title(title)
        ax.set_xticks(positions + width * (len(labels) - 1) / 2)  # Adjusting xticks position to center
        ax.set_xticklabels(categories)
        ax.legend()
        
        plt.tight_layout()
        plt.savefig(column_chart_path)
        
    # Generate Column Charts for each Motivational Bucket
    generate_column_chart_with_error_bars(labels, (row[0:3] for row in values), (row[0:3] for row in standard_errors), categories[0:3], column_paths[0], f"Extrinsic Motivations {intensive_label}")
    generate_column_chart_with_error_bars(labels, (row[3:8] for row in values), (row[3:8] for row in standard_errors), categories[3:8], column_paths[1], f"Intrinsic Motivations {intensive_label}")
    generate_column_chart_with_error_bars(labels, (row[8:13] for row in values), (row[8:13] for row in standard_errors), categories[8:13], column_paths[2], f"Social Motivations {intensive_label}")
    
    
    # Horizontal Bar Charts
    def generate_horizontal_bar_charts_with_error_bars(label, val, err, categories, N_value, total_studies, horizontal_bar_chart_path):
        
        colors = {
            "Monetary": "yellow",
            "Own-Use": "yellow",
            "Learning": "yellow",
            
            "Task-Related": "green",
            "Autonomy": "green",
            "Self-Efficacy": "green",
            "Ideology": "green",
            "Identity": "green",
            
            "Peer-Status": "orange",
            "Career-Concerns": "orange",
            "Community": "orange",
            "Reciprocity": "orange",
            "Altruism": "orange"
        }

        # Get indices to sort the values in descending order
        sorted_indices = np.argsort(val)
        sorted_vals = [val[i] for i in sorted_indices]
        sorted_errs = [err[i] * 1.96 for i in sorted_indices]
        sorted_categories = [f'{categories[i]} (N = {N_value[i]})' for i in sorted_indices]
        sorted_colors = [colors.get(cat.split()[0]) for cat in sorted_categories]

        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Plotting the bars
        bars = ax.barh(sorted_categories, sorted_vals, xerr=sorted_errs, align='center', color=sorted_colors, capsize=5)

        # Customizing the appearance of each tick label
        for bar, cat in zip(bars, sorted_categories):
            if 'Monetary' in cat:
                bar.set_edgecolor('black')
                bar.set_linewidth(1.5)

        # Update y-axis labels
        for y_label in ax.get_yticklabels():
            if 'Monetary' in y_label.get_text():
                y_label.set_fontweight('bold')
        
        ax.set_xlabel('Values')
        ax.set_title(f'{label} {intensive_label} (N = {total_studies})')
        
        # Create custom legend
        from matplotlib.lines import Line2D
        legend_elements = [
            Line2D([0], [0], color='yellow', lw=4, label='Extrinsic'),
            Line2D([0], [0], color='green', lw=4, label='Intrinsic'),
            Line2D([0], [0], color='orange', lw=4, label='Social')
        ]
        ax.legend(handles=legend_elements, loc="best")
        
        plt.tight_layout()
        plt.savefig(horizontal_bar_chart_path)
       
    
    # Generate Horizontal Bar Charts for each Platform
    for val, err, label, N_value, total_studies in zip(values, standard_errors, labels, N_values, total_number_of_studies):     
        generate_horizontal_bar_charts_with_error_bars(label, val, err, categories, N_value, total_studies, f'{label} Bar Chart.png')

    
    motivation_colors = [
        '#1f77b4',  # muted blue
        '#ff7f0e',  # safety orange
        '#2ca02c',  # cooked asparagus green
        '#d62728',  # brick red
        '#9467bd',  # muted purple
        '#8c564b',  # chestnut brown
        '#e377c2',  # raspberry yogurt pink
        '#7f7f7f',  # middle gray
        '#bcbd22',  # curry yellow-green
        '#17becf',  # blue-teal
        '#2e8b57',  # sea green
        '#ff4500',  # orangered
        '#8a2be2'   # blue violet
    ]
    
    motivations = [
        "Monetary",
        "Own-Use",
        "Learning",
        
        "Task-Related",
        "Autonomy",
        "Self-Efficacy",
        "Ideology",
        "Identity",
        
        "Peer-Status",
        "Career-Concerns",
        "Community",
        "Reciprocity",
        "Altruism"
    ]
    
    # X Y Plot of N vs Mean + SE
    def generate_xy_plot_with_error_bars(N_values_platform, total_number_of_studies_for_platform, values_platform, errors_platform, label, colors, motivations, xy_plot_path):
        fig, ax = plt.subplots(figsize=(10, 6))
        
        for N_value, value, error, color, motivation in zip(N_values_platform, values_platform, errors_platform, colors, motivations):
            ax.errorbar((N_value / total_number_of_studies_for_platform), value, yerr=error * 1.96, color=color, fmt='o', capsize=5, label=f'{motivation} (N = {int(N_value)})')
        
        ax.set_xlabel(f'Percentage of Total Number of Study Contexts (N = {int(total_number_of_studies_for_platform)})')
        ax.set_ylabel('Mean')
        ax.legend()
        ax.set_title(f'{label} - Percentage of Total Number of Study Contexts vs Mean {intensive_label}')
        plt.tight_layout()
        plt.savefig(xy_plot_path)
    
    
    # Generate XY Plots for each Platform
    for N_values_platform, total_number_of_studies_for_platform, values_platform, errors_platform, label in zip(N_values, total_number_of_studies, values, standard_errors, labels):
        generate_xy_plot_with_error_bars(N_values_platform, total_number_of_studies_for_platform, values_platform, errors_platform, label, motivation_colors, motivations, f'{label} - N XY Plot.png')

    
    plt.close('all')


# Run the Meta Analysis

def run_meta_analysis(output_directory, intensive):
    
    make_intensive_extensive_directory(META_ANALYSIS_PARENT_PATH, output_directory, intensive)
    
    CUMULATIVE_TABLE = [["Cumulative Table"]]
    
    ALL_PLATFORMS, TOTAL_NUMBER_OF_STUDIES = parse_all_tallies(
        file_path = EXCEL_SHEET_PATH,
        meta_analysis_parent_path = META_ANALYSIS_PARENT_PATH,
        names = PLATFORM_NAMES,
        intensive = intensive
    )
    
    generate_table_labels(
        cumulative_table = CUMULATIVE_TABLE,
        motivations = MOTIVATION_COLUMNS,
        platforms = PLATFORM_NAMES
    )
    
    for p in range(len(ALL_PLATFORMS)):
        
        for m in range(len(MOTIVATION_COLUMNS)):
            
            convert_to_csv(
                parsed_platform = ALL_PLATFORMS[p],
                motivation = MOTIVATION_COLUMNS[m],
                csv_path = CSV_PATH_NAMES[p],
                intensive = intensive
            )
            
            new_platform_motivation_summary_path = generate_platform_motivation_summaries(
                platform = PLATFORM_NAMES[p],
                motivation_number = m,
                motivation = MOTIVATION_COLUMNS[m][1],
                platform_path = CSV_PATH_NAMES[p],
                r_script_path = R_SCRIPT_EXECUTABLE_PATH,
                r_file_path = R_FILE_PATH
            )
            
            get_summary_statistics(
                motivation_number = m,
                platform_number = p,
                motivation = MOTIVATION_COLUMNS[m][1],
                platform = ALL_PLATFORMS[p][0],
                platform_motivation_summary_path = new_platform_motivation_summary_path,
                platform_summary_path = PLATFORM_SUMMARY_PATHS[p],
                motivation_summary_path = MOTIVATION_SUMMARY_PATHS[m],
                cumulative_table = CUMULATIVE_TABLE
            )
    
    for p in range(len(ALL_PLATFORMS)):
        
        new_platform_summary_path = generate_summary(
            summary_type = "Platforms",
            opposing_summary_type = "Motivation",
            summary = PLATFORM_NAMES[p],
            summary_path = PLATFORM_SUMMARY_PATHS[p],
            r_script_path = R_SCRIPT_EXECUTABLE_PATH,
            r_file_path = R_FILE_PATH
        )
        
        record_total_number_of_studies(
            cumulative_table = CUMULATIVE_TABLE,
            platform_number = p,
            total_number_of_studies = TOTAL_NUMBER_OF_STUDIES
        )
        
        update_table_data(
            cumulative_table = CUMULATIVE_TABLE,
            row = p+1,
            column = 0,
            summary_path = new_platform_summary_path
        )
        
    for m in range(len(MOTIVATION_COLUMNS)):
        
        new_motivation_summary_path = generate_summary(
            summary_type = "Motivations",
            opposing_summary_type = "Platform",
            summary = MOTIVATION_COLUMNS[m][1],
            summary_path = MOTIVATION_SUMMARY_PATHS[m],
            r_script_path = R_SCRIPT_EXECUTABLE_PATH,
            r_file_path = R_FILE_PATH
        )
        
        update_table_data(
            cumulative_table = CUMULATIVE_TABLE,
            row = 0,
            column = m+1,
            summary_path = new_motivation_summary_path
        )
        
    save_cumulative_table(
        cumulative_table = CUMULATIVE_TABLE,
        cumulatve_table_path = CUMULATIVE_TABLE_PATH
    )


# Run Intensive and Extensive Meta Analyses and Generate Figures
    
def main(output_directory):
    
    make_output_directory(
        meta_analysis_parent_path = META_ANALYSIS_PARENT_PATH,
        output_directory = output_directory
    )
    
    
    # Intensive Analysis (Likert)
    run_meta_analysis(
        output_directory = output_directory,
        intensive = True
    )
    
    # Extensive Analysis (Proportions, Count)
    run_meta_analysis(
        output_directory = output_directory,
        intensive = False
    )
    
    
    # Generate Intensive Plots for Meta Analysis Results
    generate_plots(
        cumulative_table_csv_path = CUMULATIVE_TABLE_PATH,
        spider_path = SPIDER_PATH,
        column_paths = COLUMN_PATHS,
        meta_analysis_parent_path = META_ANALYSIS_PARENT_PATH,
        output_directory = output_directory,
        intensive = True
    )
    
    # Generate Extensive Plots for Meta Analysis Results
    generate_plots(
        cumulative_table_csv_path = CUMULATIVE_TABLE_PATH,
        spider_path = SPIDER_PATH,
        column_paths = COLUMN_PATHS,
        meta_analysis_parent_path = META_ANALYSIS_PARENT_PATH,
        output_directory = output_directory,
        intensive = False
    )
    
    
# Final Output

RUN_META_ANALYSIS = True
OUTPUT_DIRECTORY = r"Latest Meta Analysis"

if __name__ == "__main__":
    
    if (not RUN_META_ANALYSIS):
        print("Meta Analysis is not enabled")
    
    else:
        main(OUTPUT_DIRECTORY)