import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from io import BytesIO
import matplotlib.font_manager as fm

# Load local font
def load_font(font_name):
    font_paths = {
        "Arial": "Fonts/arial.ttf",
        "Helvetica": "Fonts/Helvetica.ttf",
        "Times New Roman": "Fonts/times-new-roman.ttf",
        "Courier New": "Fonts/Courier New.ttf"
    }
    font_path = font_paths.get(font_name, "Fonts/arial.ttf")  # Default to Arial if not found
    return fm.FontProperties(fname=font_path)

# Function to calculate y positions based on bar widths
def calculate_y_positions(widths):
    positions = np.zeros(len(widths))
    positions[0] = widths[0] / 2
    for i in range(1, len(widths)):
        positions[i] = positions[i - 1] + 0.5 * (widths[i] + widths[i - 1])
    return positions

# Convert RGB values to 0-1 range
def normalize_rgb(color_list):
    return [(r / 255, g / 255, b / 255) for r, g, b in color_list]

# Function to create custom legend handles
def create_legend_handles(dataframe):
    return [mpatches.Patch(color=row[1], label=row[0]) for row in dataframe.itertuples(index=False)]

# MACC plot function with dynamic legend position
# MACC plot function with dynamic legend position
def macc_plot(df, Column1, Column2, Column3, Column4, colors, font_property, line_thickness, plot_width, plot_height,
              show_technology_labels, show_axis_titles, show_chart_title, chart_title,
              x_axis_title, y_axis_title, axis_label_font_size, axis_title_font_size,
              chart_title_font_size, technology_label_font_size, legend_font_size):
    
    df = df[[Column1, Column2, Column3, Column4, 'Label?']].dropna().sort_values(by=Column4, ascending=True)
    category = df[Column1].values
    technology = df[Column2].values
    height = df[Column4].values
    width = np.abs(df[Column3].values)
    y_pos = calculate_y_positions(width)
    
    color_df = pd.DataFrame({Column1: df[Column1].unique(), 'Colour': colors})
    df = pd.merge(df, color_df, on=Column1, how='left')
    colors = df['Colour']
    
    plt.figure(figsize=(plot_width, plot_height))
    bars = plt.bar(y_pos, height, width=width, color=colors, edgecolor='black', linewidth=line_thickness)
    
    max_label_height = 0
    if show_technology_labels:
        label_offsets = np.linspace(-0.5, 0.5, len(bars))
        for i, bar in enumerate(bars):
            if df.iloc[i]['Label?'] == 'Yes':
                label_x = bar.get_x() + bar.get_width() / 2
                label_y = bar.get_height() + 0.1 * np.abs(height.max()) if height[i] >= 0 else 0 + 0.1 * np.abs(height.max())
                va = 'bottom'
                arrow_start = bar.get_height() if height[i] >= 0 else 0

                plt.annotate(
                    technology[i],
                    xy=(label_x, arrow_start),
                    xytext=(label_x + label_offsets[i], label_y + 0.2 * np.abs(height.max())),
                    textcoords='data',
                    ha='center',
                    va=va,
                    fontsize=technology_label_font_size,
                    rotation=90,
                    fontproperties=font_property,
                    arrowprops=dict(arrowstyle="->", color='gray', lw=1.5)
                )
                max_label_height = max(max_label_height, label_y + 0.2 * np.abs(height.max()))

    plt.axhline(0, color='black', linewidth=line_thickness)
    
    if show_chart_title:
        plt.title(chart_title, fontsize=chart_title_font_size, fontproperties=font_property)
    
    if show_axis_titles:
        plt.ylabel(y_axis_title, fontsize=axis_title_font_size, fontproperties=font_property)
        plt.xlabel(x_axis_title, fontsize=axis_title_font_size, fontproperties=font_property)
    
    plt.xlim(y_pos.min() - 0.5 * width[0], y_pos.max() + 0.5 * width[-1])
    plt.ylim(height.min() - 0.1 * np.abs(height.min()), height.max() + 0.5 * np.abs(height.max()))
    
    legend_handles = create_legend_handles(color_df)
    legend_y_position = 1.0
    if show_technology_labels and max_label_height > height.max():
        legend_y_position = 1.0 + (max_label_height - height.max()) / (plot_height * 100)
    
    plt.legend(handles=legend_handles, fontsize=legend_font_size, loc='upper left', bbox_to_anchor=(0, legend_y_position))
    
    ax = plt.gca()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    # Apply font size to axis labels
    ax.xaxis.set_tick_params(labelsize=axis_label_font_size)
    ax.yaxis.set_tick_params(labelsize=axis_label_font_size)
    
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
    df = pd.read_excel(uploaded_file, sheet_name="MACC Data")

    # Tabs for different views
    tab1, tab2 = st.tabs(["Plot", "Data Table"])

    with tab1:
        # Sidebar for customization
        st.sidebar.header('Plot Customization')
        font_style = st.sidebar.selectbox('Font Style', options=['Arial', 'Helvetica', 'Times New Roman', 'Courier New'])
        font_property = load_font(font_style)
        line_thickness = st.sidebar.slider('Line Thickness', min_value=0.5, max_value=5.0, value=1.0, step=0.1)
        plot_width = st.sidebar.slider('Plot Width', min_value=4.0, max_value=20.0, value=15.0, step=0.5)
        plot_height = st.sidebar.slider('Plot Height', min_value=4.0, max_value=15.0, value=10.0, step=0.5)

        # Color customization
        color_scheme_option = st.sidebar.selectbox('Color Scheme', options=['Default', 'Alternative', 'Custom'])
        default_colors = normalize_rgb([
            (47, 182, 255), (43, 134, 199), (175, 194, 54),
            (225, 39, 141), (255, 212, 0), (233, 90, 26), (0, 150, 149)
        ])

        if color_scheme_option == 'Alternative':
            alternative_colors = normalize_rgb([
                (255, 99, 132), (54, 162, 235), (255, 206, 86), (75, 192, 192),
                (153, 102, 255), (255, 159, 64), (199, 199, 199)
            ])
            colors = alternative_colors
        elif color_scheme_option == 'Custom':
            custom_colors = []
            for idx, category in enumerate(df['Category'].unique()):
                default_color = '#{:02x}{:02x}{:02x}'.format(*[int(c*255) for c in default_colors[idx % len(default_colors)]])
                color = st.sidebar.color_picker(f"Pick a color for {category}", default_color)
                custom_colors.append(tuple(int(color[i:i+2], 16) for i in (1, 3, 5)))
            colors = normalize_rgb(custom_colors)
        else:
            colors = default_colors

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
        plt = macc_plot(df, 'Category', 'Technology Option', 'ktCO2e', '£/tCO2e', colors, font_property, line_thickness, plot_width, plot_height,
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
