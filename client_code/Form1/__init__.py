from ._anvil_designer import Form1Template
from anvil import *
import anvil.tables as tables
import anvil.tables.query as q
from anvil.tables import app_tables
import anvil.server


class Form1(Form1Template):
  def __init__(self, **properties):
    # Set Form properties and Data Bindings.
    self.init_components(**properties)

  def outlined_button_1_click(self, **event_args):
    """This method is called when the button is clicked"""
    master_mapping_file = self.master_mapping_loader.file
    input_file = self.input_file_loader.file  # The file loaded by the FileLoader
    sheetname = self.text_box_1.text
    output_file_name = self.text_box_3.text 
    low_confidence_score = self.text_box_4.text
    xmed_confidence_score = self.text_box_5.text

    try:
      print("attempting to access python server...")
      output_file = anvil.server.call(
          'standardize',
          master_mapping_file,
          input_file, 
          sheetname, 
          output_file_name,  
          low_confidence_score,
          xmed_confidence_score
      )
      
      download(output_file)
      # Handle the returned output file (provide it for download)
      
    except Exception as e:
      # Handle errors
      return("error", e)

    print('now setting download link')
    print(self.download_link_info)
    self.download_link_info.visible = True
    self.download_link_info.content = "Output file was successfully downloaded to your computer. If you need to redownload the file, please click this link: "
    print('download link info should be visible')
    self.download_link.visible = True
    self.download_link.text = "DOWNLOAD OUTPUT FILE"
    self.download_link.file = output_file

  def master_mapping_loader_change(self, file, **event_args):
    """This method is called when a new file is loaded into this FileLoader"""
    
    if file:
      self.mm_file_name.text = file.name

  def input_file_loader_change(self, file, **event_args):
    """This method is called when a new file is loaded into this FileLoader"""
    if file:
      self.input_file_name.text = file.name

  def download_link_click(self, **event_args):
    """This method is called when the link is clicked"""
    download(self.download_link.file)

  def button_1_click(self, **event_args):
    """This method is called when the button is clicked"""
    print('button 1 clicked')
    ans = anvil.server.call("display_name", "Collin")
    print('ans was recieved')
    print(ans)

  def text_box_4_change(self, **event_args):
    """This method is called when the text in this text box is edited"""
    pass

  def text_box_4_pressed_enter(self, **event_args):
    """This method is called when the user presses Enter in this text box"""
    low_score = float(self.text_box_4.text)
    print(f'low score = {low_score}')

    if low_score < 0 or low_score > 1:
        print('ERROR, input must be between 0 and 1')

  def project_info_link_click(self, **event_args):
    """This method is called when the link is clicked"""
    # download(self.project_info_link.file)
    pass

  
