import gradio as gr
import plotly.graph_objects as go

def plot_example():
    fig = go.Figure()
    fig.add_scatter(y=[1, 2, 3])
    return fig

demo = gr.Interface(fn=plot_example, inputs=[], outputs=gr.Plot())
demo.launch()