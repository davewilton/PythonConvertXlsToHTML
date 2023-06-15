# pylint: disable=locally-disabled

import re
import markdown
import pandas as pd


def convert_excel_to_html(excel_file, output_file):
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Iterate over each cell in the DataFrame
    for col in df.columns:
        for i, cell in enumerate(df[col]):
            if pd.notna(cell):
                # Find links in the cell's text formatted with Markdown syntax
                links = re.findall(r'\[(.*?)\]\((.*?)\)', str(cell))
                for link_text, link_url in links:
                    # Replace each Markdown link with an HTML link
                    html_link = f'<a href="{link_url}">{link_text}</a>'
                    cell = cell.replace(f'[{link_text}]({link_url})', html_link)
                df.at[i, col] = cell
                # Convert Markdown to HTML for remaining Markdown-formatted text
                cell = markdown.markdown(cell)
                cell = cell.replace('\n', '')
                df.at[i, col] = cell

    # Convert the DataFrame to HTML
    html_table = df.to_html(escape=False, index=False)

    # Write HTML to output file
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(html_table)


convert_excel_to_html('CottonRecommendations.xlsx', 'output.html')
