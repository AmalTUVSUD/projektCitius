from document_processor.doc_handler import DocHandler


# Initialize
doc_handler = DocHandler()

# Open file
doc =doc_handler.open_doc("C:/Users/smaou-am/projektCitius/input/Test_TRF.doc")
print("Document opened:", doc_handler.doc is not None)

# Count tables
doc_handler.print_first_3_tables()
# Save changes
doc_handler.save_doc()

# Clean up
doc_handler.close_doc()