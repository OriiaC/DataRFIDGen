using System;
using System.Windows.Forms;


namespace Data_RFID_Gen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string inicialTag = "1700120";

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
            int iteraciones = Int32.Parse(textBox2.Text);
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
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
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
            textBox3.SelectAll();
            textBox3.Copy();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox4.SelectAll();
            textBox4.Copy();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox5.SelectAll();
            textBox5.Copy();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBox6.SelectAll();
            textBox6.Copy();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && (e.KeyChar != (char)Keys.Back)){
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsNumber(e.KeyChar) && (e.KeyChar != (char)Keys.Back))
            {
                e.Handled = true;
            }
        }

        static string extraerCaracteres(string cadena, int numeroCaracteres)
        {
            int tam_cadena = cadena.Length;
            return cadena.Substring((tam_cadena - numeroCaracteres), numeroCaracteres);
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
    }
}
