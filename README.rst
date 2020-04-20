##############
docx-equation
##############

convert equation inside word(.docx) to latex

These is a demo file named **equation.docx**. 
You can use the script downside to convert it into html file.
The equation inside **equation.docx** will be transformed into LaTeX format.


********
Usuage
********


.. code:: python

    from docx_equation.docx import convert_to_html
    convert_to_html('equation.docx')


.. code:: html

    <p>${\left(x+a\right)}^{n}={\sum }_{k=0}^{n}\left(\begin{array}{c}n\\ k\end{array}\right){x}^{k}{a}^{n-k}$</p>
