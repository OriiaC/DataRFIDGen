using System;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using objExcel = Microsoft.Office.Interop.Excel;

namespace Data_RFID_Gen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox1.Text.Trim() == string.Empty || textBox2.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Por favor rellene los campos obligatorios (*)");
                return;
            }
            long inicial = long.Parse(textBox1.Text);
            if (radioButton1.Checked == true)
            {
                if (inicial >= 17001202000000 && inicial < 17001203000000)
                {
                    string epc = "33140A607000008003";
                    limpiaCajas();
                    agregaEncabezados();
                    llenaCajas(inicial, epc);
                }
                else
                {
                    MessageBox.Show("El tag inicial de ME es incorrecto");
                }
            }
            else if (radioButton2.Checked == true)
            {
                if (inicial >= 17001203000000)
                {
                    string epc = "33140A60700000C003";
                    limpiaCajas();
                    agregaEncabezados();
                    llenaCajas(inicial, epc);
                }
                else
                {
                    MessageBox.Show("El tag inicial de TI es incorrecto");
                }
            }

        }

        string creaTexto(long tag)
            //concatena el prefijo con el tag
        {
            if (radioButton1.Checked == true)
            {
                string texto = "ME-" + tag;
                return texto;
            }
            else
            {
                string texto = "TI-" + tag;
                return texto;
            }            
        }

        void limpiaCajas()
        {
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            EPC.DataGridView.Rows.Clear();
            BarCode.DataGridView.Rows.Clear();
            Texto.DataGridView.Rows.Clear();
            EPCCompleto.DataGridView.Rows.Clear();


        }

        void agregaEncabezados()
        {
            textBox3.Text = "EPC" + Environment.NewLine;
            textBox4.Text = "Barcode" + Environment.NewLine;
            textBox5.Text = "Texto" + Environment.NewLine;
            textBox6.Text = "EPC,Barcode,Texto" + Environment.NewLine;
        }

        void llenaCajas(long inicialTag, string inicialEPC)
        {
            //captura la cantidad de etiquetas que va a generar y las guarda en la variable iteraciones
            int iteraciones = Int32.Parse(textBox2.Text);
            //agrega los datos a los textbox y alos datagridview
            for (int i = 0; i <= iteraciones; i++)
            {
                long actualTag = inicialTag + i;
                string stringTag = actualTag.ToString();
                string ultimos6 = extraerCaracteres(stringTag,6);
                string finalEPC = inicialEPC + ultimos6;
                textBox3.AppendText(finalEPC + Environment.NewLine);
                textBox4.AppendText(actualTag.ToString() + Environment.NewLine);
                textBox5.AppendText(creaTexto(actualTag) + Environment.NewLine);
                textBox6.AppendText(finalEPC+"," + actualTag.ToString() + "," + creaTexto(actualTag) + Environment.NewLine);

                dataGridView1.Rows.Add(finalEPC);
                dtgvBarCode.Rows.Add(actualTag.ToString());
                dtgvTexto.Rows.Add(creaTexto(actualTag));
                dtgvTotal.Rows.Add();
                dtgvTotal.Rows[i].Cells[0].Value = finalEPC;
                dtgvTotal.Rows[i].Cells[1].Value = actualTag;
                dtgvTotal.Rows[i].Cells[2].Value = creaTexto(actualTag);

            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //agrega los campos de texto de encabezado en los textbox
            agregaEncabezados();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            limpiaCajas();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //Selecciona y copia la data generada en los text box
            //textBox3.SelectAll();
            //textBox3.Copy();
            //activa el modo para copiar el DataGridView
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            //Selecciona todo el DataGridView
            dataGridView1.SelectAll();
            //Copia el DataGridView
            DataObject dataObj = dataGridView1.GetClipboardContent();
            Clipboard.SetDataObject(dataObj, true);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Selecciona y copia la data generada en los text box
            //textBox4.SelectAll();
            //textBox4.Copy();
            //Selecciona y copia la data generada en los text box
            //textBox3.SelectAll();
            //textBox3.Copy();
            //activa el modo para copiar el DataGridView
            dtgvBarCode.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            //Selecciona todo el DataGridView
            dtgvBarCode.SelectAll();
            //Copia el DataGridView
            DataObject dataObj = dtgvBarCode.GetClipboardContent();
            Clipboard.SetDataObject(dataObj, true);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //Selecciona y copia la data generada en los text box
            //textBox5.SelectAll();
            //textBox5.Copy();
            //Selecciona y copia la data generada en los text box
            //textBox3.SelectAll();
            //textBox3.Copy();
            //activa el modo para copiar el DataGridView
            dtgvTexto.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            //Selecciona todo el DataGridView
            dtgvTexto.SelectAll();
            //Copia el DataGridView
            DataObject dataObj = dtgvTexto.GetClipboardContent();
            Clipboard.SetDataObject(dataObj, true);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //Selecciona y copia la data generada en los text box
            textBox6.SelectAll();
            textBox6.Copy();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //no deja escribir caracteres alfabeticos en los textbox numericos
            if (!char.IsNumber(e.KeyChar) && (e.KeyChar != (char)Keys.Back)){
                e.Handled = true;
            }
        }
        
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //no deja escribir caracteres alfabeticos en los textbox numericos
            if (!char.IsNumber(e.KeyChar) && (e.KeyChar != (char)Keys.Back))
            {
                e.Handled = true;
            }
        }
        //Toma los ultimo x caracteres indicados
        static string extraerCaracteres(string cadena, int numeroCaracteres)
        {
            int tam_cadena = cadena.Length;
            return cadena.Substring((tam_cadena - numeroCaracteres), numeroCaracteres);
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            objExcel.Application objAplicacion= new objExcel.Application();
            Workbook objLibro = objAplicacion.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet objHoja = (Worksheet)objAplicacion.ActiveSheet;
            //Se crea el encabezado
            objAplicacion.Visible = false;
            objHoja.Cells[1, 1] = "EPC";
            objHoja.Cells[1, 2] = "Barcode";
            objHoja.Cells[1, 3] = "Texto";
            //captura la cantidad de etiquetas a generar
            int iteraciones = Int32.Parse(textBox2.Text);
            //itera y llena el excel
            for (int i = 2; i <= iteraciones+1; i++)
            {
                long Tag = long.Parse(textBox1.Text) + i-2;
                string stringTag = Tag.ToString();
                string ultimosEPC=extraerCaracteres(stringTag, 6);
                string primerosEPC = "33140A607000008003";
                objHoja.Cells[i, 1] = primerosEPC + ultimosEPC;
                objHoja.Cells[i, 2] = Tag;
                objHoja.Cells[i, 3] = "ME-"+Tag;
            }
            try
            {
                objLibro.SaveAs(ruta + "\\Data_RFID_Gen_BdB.xlsx");
                objLibro.Close();
                objAplicacion.Quit();
            }catch (Exception ex)
            {
                MessageBox.Show("No se creara el documento Excel.\r\nError:" + ex);
            }
            
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            //activa el modo para copiar el DataGridView
            dtgvTotal.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            //Selecciona todo el DataGridView
            dtgvTotal.SelectAll();
            //Copia el DataGridView
            DataObject dataObj = dtgvTotal.GetClipboardContent();
            Clipboard.SetDataObject(dataObj, true);
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void dtgvTotal_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //Selecciona y copia la data generada en los text box
            //textBox3.SelectAll();
            //textBox3.Copy();
            //activa el modo para copiar el DataGridView
            dtgvTotal.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            //Selecciona todo el DataGridView
            dtgvTotal.SelectAll();
            //Copia el DataGridView
            DataObject dataObj = dtgvTotal.GetClipboardContent();
            Clipboard.SetDataObject(dataObj, true);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
