import asyncio
import json
from word_document_server.tools.document_tools import create_document
from word_document_server.tools.content_tools import (
    add_paragraph, add_table, get_document_blocks, read_document_block,
    modify_document_block, delete_document_block
)

async def test():
    filename = "test_unified.docx"
    print("Creating document...")
    await create_document(filename)

    print("Adding paragraph...")
    await add_paragraph(filename, "First paragraph")

    print("Adding table...")
    await add_table(filename, 2, 2, [["A1", "B1"], ["A2", "B2"]])

    print("Adding second paragraph...")
    await add_paragraph(filename, "Second paragraph")

    print("\n--- Blocks ---")
    blocks = await get_document_blocks(filename)
    print(blocks)

    if "Failed" in blocks:
        return

    print("\n--- Read Block 1 (Table) ---")
    tbl = await read_document_block(filename, 1)
    print(tbl)

    print("\n--- Modify Block 1 (Table) ---")
    await modify_document_block(filename, 1, table_data=[["C1", "C2"], ["C3", "C4"]])
    tbl_mod = await read_document_block(filename, 1)
    print(tbl_mod)

    print("\n--- Modify Block 0 (Paragraph) ---")
    await modify_document_block(filename, 0, paragraph_text="Modified First paragraph", bold=True, color="FF0000")
    para_mod = await read_document_block(filename, 0)
    print(para_mod)

    print("\n--- Delete Block 2 (Second Paragraph) ---")
    await delete_document_block(filename, 2)
    
    print("\n--- Final Blocks ---")
    blocks_final = await get_document_blocks(filename)
    print(blocks_final)

if __name__ == "__main__":
    asyncio.run(test())
