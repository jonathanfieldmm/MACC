import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from io import BytesIO

# Function for position depending on width
def y_pos1(width):
    a = np.zeros(len(width))
    a[0] = width[0] / 2
    for i in range(1, len(width)):
        a[i] = a[i - 1] + 0.5 * width[i] + 0.5 * width[i - 1]
    return a

# Convert RGB values to 0-1 range
def normalise_rgb(colour_list):
    return [(r / 255, g / 255, b / 255) for r, g, b in colour_list]

# Function to customize legend
def leg_handles(dataframe):
    a = []
    for i in range(len(dataframe)):
        patch = mpatches.Patch(color=dataframe.iloc[i, 1], label=dataframe.iloc[i, 0])
        a.append(patch)
    return a

# MACC plot function with dynamic legend position
def macc_function(df, Column1, Column2, Column3, Column4, colours, font_style, line_thickness, plot_width, plot_height,
                  show_technology_labels, show_axis_titles, show_chart_title, chart_title,
                  x_axis_title, y_axis_title, axis_label_font_size, axis_title_font_size,
                  chart_title_font_size, technology_label_font_size, legend_font_size):
    df1 = df[[Column1, Column2, Column3, Column4, 'Label?']].dropna().sort_values(by=Column4, ascending=True)
    category = df1[Column1].values.squeeze()  # Categories
    technology = df1[Column2].values.squeeze()  # Technology options
    height = df1[Column4].values.squeeze()  # Cost-effectiveness
    width = np.array([abs(i) if i < 0 else i for i in df1[Column3].values.squeeze()])  # Emission reduction
    y_pos = y_pos1(width)
    
    d = {Column1: df1[Column1].unique(), 'Colour': colours}
    df2 = pd.DataFrame(data=d)
    df1 = pd.merge(df1, df2, on=Column1, how='left')
    colours = df1['Colour']
    
    plt.figure(figsize=(plot_width, plot_height))
    bars = plt.bar(y_pos, height, width=width, color=colours, edgecolor='black', linewidth=line_thickness)
    
    if show_technology_labels:
        label_offsets = np.linspace(-0.5, 0.5, len(bars))  # Create varying offsets for labels
        max_label_height = 0  # To track the maximum label height
        for i, bar in enumerate(bars):
            if df1.iloc[i]['Label?'] == 'Yes':  # Check if this option should be labeled
                label_x = bar.get_x() + bar.get_width() / 2  # Midpoint for both positive and negative bars
                if height[i] >= 0:
                    label_y = bar.get_height() + 0.1 * abs(height.max())  # Position label just above the bar
                    va = 'bottom'
                    arrow_start = bar.get_height()
                else:
                    label_y = 0 + 0.1 * abs(height.max())  # Position label just above the zero line for negative bars
                    va = 'bottom'
                    arrow_start = 0  # Start arrow at the baseline (zero)

                # Annotate with vertical text and leader lines
                plt.annotate(
                    technology[i], 
                    xy=(label_x, arrow_start), 
                    xytext=(label_x + label_offsets[i], label_y + 0.2 * abs(height.max())),
                    textcoords='data', 
                    ha='center', 
                    va=va, 
                    fontsize=technology_label_font_size,
                    rotation=90, 
                    fontname=font_style,
                    arrowprops=dict(arrowstyle="->", color='gray', lw=1.5)
                )
                max_label_height = max(max_label_height, label_y + 0.2 * abs(height.max()))

    plt.axhline(0, color='black', linewidth=line_thickness)
    
    if show_chart_title:
        plt.title(chart_title, fontsize=chart_title_font_size, fontname=font_style)
    
    if show_axis_titles:
        plt.ylabel(y_axis_title, fontname=font_style, fontsize=axis_title_font_size)
        plt.xlabel(x_axis_title, fontname=font_style, fontsize=axis_title_font_size)
    
    plt.xlim(y_pos.min() - 0.5 * width[0], y_pos.max() + 0.5 * width[-1])
    plt.ylim(height.min() - 0.1 * abs(height.min()), height.max() + 0.5 * abs(height.max()))
    
    legend_handles = leg_handles(df2)
    
    # Adjust the legend position based on max label height
    legend_y_position = 1.0  # Start at the top
    if show_technology_labels and max_label_height > height.max():
        legend_y_position = 1.0 + (max_label_height - height.max()) / (plot_height * 100)
    
    plt.legend(handles=legend_handles, fontsize=legend_font_size, loc='upper left', bbox_to_anchor=(0, legend_y_position))
    
    ax = plt.gca()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    return plt

# Streamlit app setup
st.title('MACC Plot Viewer')

st.write("""
### Instructions:
1. The uploaded Excel file should have a sheet named **MACC Data**.
2. The sheet should contain the following columns:
   - **Category**: Represents the different categories (e.g., technology options).
   - **Technology Option**: Details of the technology or intervention.
   - **ktCO2e**: Emission reduction potential in kilotonnes of CO2 equivalent.
   - **£/tCO2e**: Cost-effectiveness in GBP per tonne of CO2 equivalent.
   - **Label?**: A column indicating whether the technology should be labeled ("Yes" for label, "No" for no label).
3. Ensure there are no empty rows or irrelevant data outside these columns.
4. Customize your plot using the options in the sidebar. You can select the color scheme, font style, line thickness, plot dimensions, and more.
5. After customization, you can download the plot as a PNG file using the download button below the plot.
""")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
if uploaded_file is not None:
    sheet_name = "MACC Data"
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    # Tabs for different views
    tab1, tab2 = st.tabs(["Plot", "Data Table"])

    with tab1:
        # Sidebar for customization
        st.sidebar.header('Plot Customization')
        font_style = st.sidebar.selectbox('Font Style', options=['Arial', 'Helvetica', 'Times New Roman', 'Courier New'])
        line_thickness = st.sidebar.slider('Line Thickness', min_value=0.5, max_value=5.0, value=1.0, step=0.1)
        plot_width = st.sidebar.slider('Plot Width', min_value=4.0, max_value=20.0, value=15.0, step=0.5)
        plot_height = st.sidebar.slider('Plot Height', min_value=4.0, max_value=15.0, value=10.0, step=0.5)

        # Color customization
        color_scheme_option = st.sidebar.selectbox('Color Scheme', options=['Default', 'Alternative', 'Custom'])

        default_colours = normalise_rgb([
            (47, 182, 255), (43, 134, 199), (175, 194, 54),
            (225, 39, 141), (255, 212, 0), (233, 90, 26), (0, 150, 149)
        ])

        if color_scheme_option == 'Alternative':
            alternative_colours = normalise_rgb([
                (255, 99, 132), (54, 162, 235), (255, 206, 86), (75, 192, 192),
                (153, 102, 255), (255, 159, 64), (199, 199, 199)
            ])
            colours = alternative_colours
        elif color_scheme_option == 'Custom':
            custom_colours = []
            for idx, category in enumerate(df['Category'].unique()):
                default_color = '#{:02x}{:02x}{:02x}'.format(*[int(c*255) for c in default_colours[idx % len(default_colours)]])
                color = st.sidebar.color_picker(f"Pick a color for {category}", default_color)
                custom_colours.append(tuple(int(color[i:i+2], 16) for i in (1, 3, 5)))
            colours = normalise_rgb(custom_colours)
        else:
            colours = default_colours

        # Label settings
        show_technology_labels = st.sidebar.checkbox('Show Technology Labels', value=True)
        technology_label_font_size = st.sidebar.slider('Technology Label Font Size', min_value=6, max_value=20, value=10)
        show_axis_titles = st.sidebar.checkbox('Show Axis Titles', value=True)
        axis_label_font_size = st.sidebar.slider('Axis Label Font Size', min_value=6, max_value=20, value=10)
        axis_title_font_size = st.sidebar.slider('Axis Title Font Size', min_value=6, max_value=20, value=12)

        # Title settings
        show_chart_title = st.sidebar.checkbox('Show Chart Title', value=True)
        chart_title_font_size = st.sidebar.slider('Chart Title Font Size', min_value=6, max_value=20, value=14)
        chart_title = st.sidebar.text_input('Chart Title', value='MACC')
        x_axis_title = st.sidebar.text_input('X-axis Title', value='kt CO$_2$e')
        y_axis_title = st.sidebar.text_input('Y-axis Title', value='£/tCO$_2$e')

        # Legend font size customization
        legend_font_size = st.sidebar.slider('Legend Font Size', min_value=6, max_value=20, value=10)

        # Plotting
        plt = macc_function(df, 'Category', 'Technology Option', 'ktCO2e', '£/tCO2e', colours, font_style, line_thickness, plot_width, plot_height,
                            show_technology_labels, show_axis_titles, show_chart_title, chart_title,
                            x_axis_title, y_axis_title, axis_label_font_size, axis_title_font_size,
                            chart_title_font_size, technology_label_font_size, legend_font_size)
        st.pyplot(plt)

        # Add a download button
        buf = BytesIO()
        plt.savefig(buf, format="png")
        buf.seek(0)
        st.download_button(
            label="Download plot as PNG",
            data=buf,
            file_name="MACC_plot.png",
            mime="image/png"
        )

    with tab2:
        st.write("### Data Table")
        st.dataframe(df)
