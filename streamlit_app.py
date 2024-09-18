import streamlit as st
import pandas as pd
import requests
import re

def main():
    st.title("Broken Link Checker for Excel Files")

    # File uploader to accept Excel files
    uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        # Read the uploaded Excel file into a DataFrame
        df = pd.read_excel(uploaded_file, header=None)
        st.write("File uploaded successfully.")

        # Flatten the DataFrame to get all cell values with their positions
        cells = df.stack().reset_index()
        cells.columns = ['row', 'col', 'value']

        # Regular expression pattern to find URLs
        url_pattern = re.compile(r'(https?://[^\s]+)')

        # List to hold all URLs with their positions
        urls = []

        # Iterate over all cells to find URLs
        for index, row in cells.iterrows():
            cell_value = str(row['value'])
            matches = url_pattern.findall(cell_value)
            for url in matches:
                urls.append({'row': row['row'], 'col': row['col'], 'url': url})

        total_links = len(urls)
        if total_links == 0:
            st.write("No URLs found in the Excel file.")
        else:
            st.write(f"Found {total_links} URLs in the Excel file.")

            broken_links = []
            checked_links = 0

            # Initialize progress bar
            progress_bar = st.progress(0)

            # Check each URL to see if it's broken
            for idx, url_info in enumerate(urls):
                url = url_info['url']
                try:
                    # Send a HEAD request to the URL
                    response = requests.head(url, allow_redirects=True, timeout=5)
                    # Consider status codes >= 400 as broken
                    if response.status_code >= 400:
                        broken_links.append(url_info)
                except requests.RequestException:
                    # Any exception is considered as broken
                    broken_links.append(url_info)

                checked_links += 1
                progress_bar.progress(checked_links / total_links)

            # Display broken links
            if broken_links:
                st.write("Broken links found:")
                broken_links_df = pd.DataFrame(broken_links)
                # Adjust row and column indices to be 1-based instead of 0-based
                broken_links_df['row'] = broken_links_df['row'] + 1
                broken_links_df['col'] = broken_links_df['col'] + 1
                st.table(broken_links_df)
            else:
                st.write("No broken links found.")

if __name__ == "__main__":
    main()
