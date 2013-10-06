require 'docx'

SRC_FILE = 'template.docx'
DST_FILE = 'ready.docx'

# Note, in second element 'COMMENTS' is absent. Original stub in template file will be remainded
hash = {'NAME' => 'Ann', 'ELEMENTS' => [
    { 'NAME' => 'World history', 'INDEX' => '1','NUMBER' => '3', 'COMMENTS' => "Yeah, I'm so stupid in it"},
    {'INDEX' => '2', 'NAME' => 'Russian language', 'NUMBER' => '5'}
]}

w = Docx::DocxHandler.open(SRC_FILE)
w.insert(hash)
w.save(DST_FILE)
