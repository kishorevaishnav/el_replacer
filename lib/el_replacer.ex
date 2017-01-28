defmodule ElReplacer do
  def start(xlsxfile, xmlfile, path) do
    # path = "/Users/kishore/proj/SNV/VRT/XMLCopy/"
    # xlsxfile = "/Users/kishore/Downloads/vrt_narra_updated.xlsx"
    # xmlfile= "debit_note.xml"

    # Read the XLSX and get number of rows present.
    IO.puts "Reading XLS #{xlsxfile}"
    {:ok, table_id} = Xlsxir.multi_extract(xlsxfile, 0)
    total_rows = Xlsxir.get_multi_info(table_id, :rows)
    IO.puts "Total no. of rows #{total_rows}"

    # Read the file now.
    {:ok, body} = File.read("#{path}#{xmlfile}")
    IO.puts "Reading #{xmlfile}"

    # Need to use UTF16 as the file read is a UTF16.
    file_data = body |> :unicode.characters_to_binary({:utf16, :little})
    File.close body

    # Replace the unwanted line ending characters.
    file_data = Regex.replace(~r/&#13;&#10;/, file_data,  "")

    # Replace now based on the Excel (XLSX).
    file_data = Enum.reduce(2..total_rows, file_data, fn(x, acc) ->
        [file, narra, repl, _ ] = Xlsxir.get_row(table_id, x)
        if file == xmlfile do
          IO.puts "#{file} row no. #{x}"
          if is_nil(narra)== false and is_nil(repl) == false do
            narra = Regex.escape(narra)
            Regex.replace(~r/#{narra}/, acc,  repl)
          else
            acc
          end
        else
          acc
        end
      end
    )
    Xlsxir.close

    # Write back to a new file.
    {:ok, file} = File.open "new_#{xmlfile}", [:write]
    IO.binwrite file, file_data
    File.close file
  end
end

# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","contra.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","credit_note.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","credit_note_0.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","debit_note.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","journal.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","online_cr.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","online_sales.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","payment.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","purchase.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","purchase_5.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","purchase_cst.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","receipt.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","sale_0.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","sales.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
# ElReplacer.start("/Users/kishore/Downloads/vrt_narra_updated.xlsx","stock_journal.xml", "/Users/kishore/proj/SNV/VRT/XMLCopy/")
