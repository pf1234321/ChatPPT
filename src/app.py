# app.py
import gradio as gr

from chatppt_processor import process_user_input


#
# print(fsspec.__version__)
def chatbot_interface(user_input):
    output_pptx_path = process_user_input(user_input)
    return output_pptx_path


iface = gr.Interface(
    fn=chatbot_interface,
    inputs="text",
    outputs="file",
)

if __name__ == "__main__":
    iface.launch()
