import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.colors import qualitative

# Function to position the bars based on their width
def y_pos1(width):
    a = np.zeros(len(width))
    a[0] = width[0] / 2
    for i in range(1, len(width)):
        a[i] = a[i - 1] + 0.5 * width[i] + 0.5 * width[i - 1]
    return a

# Convert hex or RGB to Plotly compatible format
def normalize_color(colour_list):
    return [f'rgb({r}, {g}, {b})' for r, g, b in colour_list]

# Function to create the plot
def macc_plotly(df, Column1, Column2, Column3, Column4, colours):
    df1 = df[[Column1, Column2, Column3, Column4]].dropna().sort_values(by=Column4, ascending=True)
    technology = df1[Column2].values
    height = df1[Column4].values
    width = np.array([abs(i) for i in df1[Column3].values])
    y_positions = y_pos1(width)

    fig = go.Figure()
    for i in range(len(df1)):
        fig.add_trace(go.Bar(
            x=[y_positions[i]],
            y=[height[i]],
            width=[width[i]],
            name=technology[i],
            marker_color=colours[i % len(colours)],
            offset=0,
            customdata=[technology[i]],
            hovertemplate="Technology: %{customdata[0]}<br>Reduction: %{width}<br>Cost: %{y}",
        ))

    fig.update_layout(
        title='MACC Chart',
        xaxis_title='Emission Reduction (ktCO2e)',
        yaxis_title='Cost (£/tCO2e)',
        bargap=0.2
    )
    return fig

# Streamlit app setup
st.title('Interactive MACC Plot Viewer')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name='MACC Data')

    colours = [(255, 99, 132), (54, 162, 235), (255, 206, 86), (75, 192, 192),
               (153, 102, 255), (255, 159, 64), (199, 199, 199)]
    colours = normalize_color(colours)

    fig = macc_plotly(df, 'Category', 'Technology Option', 'ktCO2e', '£/tCO2e', colours)
    st.plotly_chart(fig, use_container_width=True)

    # Add a button to download the plot as a static image
    fig.write_image("MACC_plot.png")
    with open("MACC_plot.png", "rb") as file:
        btn = st.download_button(
            label="Download plot as PNG",
            data=file,
            file_name="MACC_plot.png",
            mime="image/png"
        )
