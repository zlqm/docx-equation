##############
docx-equation
##############

convert equation inside word(.docx) file to latex, current only supports omml

***********
background
***********


First of all, word(.docx) file is a zip file.
After unziping a word file, you can get several documents.
The content is mainly stored at **word/document.xml**.

    .. code:: 
    
        $: unzip equation.docx
        $: tree
        ├── equation.docx
        ├── [Content_Types].xml
        ├── docProps
        │   ├── app.xml
        │   └── core.xml
        ├── equation.docx
        ├── _rels
        └── word
            ├── document.xml
            ├── fontTable.xml
            ├── _rels
            │   └── document.xml.rels
            └── styles.xml
    

There are three ways to insert a equation into word file:

    1. eq field 
    2. omml
    3. MathType 

All of them are stored as markup inside docx file.
You can easily find them at **word/document.xml**.

.. code:: xml

    <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:sSup>
            <m:e>
                <m:d>
                    <m:dPr>
                        <m:begChr m:val="("/>
                        <m:endChr m:val=")"/>
                    </m:dPr>
                    <m:e>
                        <m:r>
                            <w:rPr>
                                <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                            </w:rPr>
                            <m:t xml:space="preserve">x</m:t>
                        </m:r>
                    .......
                    </m:e>
                </m:sSup>
            </m:e>
        </m:nary>
    </m:oMath>

So we can just extract the equation markup out and convert it into LaTeX.

********
Demo
********

There is a demo word file named **equation.docx** which contains a euqation.
Install requirements and run the **demo.py** and you will get **equation.html** contains the LaTeX format equation.

.. code:: html

    <p>${\left(x+a\right)}^{n}={\sum }_{k=0}^{n}\left(\begin{array}{c}n\\ k\end{array}\right){x}^{k}{a}^{n-k}$</p>
