using System.ComponentModel;
using System.IO.Ports;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using SelectPdf;
using System.Web;
using static System.Net.Mime.MediaTypeNames;
using System.Xml.Linq;
using Syncfusion.DocIO.DLS;
using MiniSoftware;

namespace EV_Testing_report_DEMO
{

    public partial class Form1 : Form
    {
        SerialPort esp32_module;
        ReportReceive_states receive_present_state;

        /*----------- Testing result ---------------------*/
        State_Transition_Test? result_State_A_to_B = new State_Transition_Test { };
        State_Transition_Test? result_State_B_to_C = new State_Transition_Test { };
        State_Transition_Test? result_State_C_to_B = new State_Transition_Test { };
        State_Transition_Test? result_State_B_to_D = new State_Transition_Test { };
        Diode_Test? result_diode_test = new Diode_Test { };
        RCD0? result_rcd_test = new RCD0 { };
        Insulation_Test? Insulation_Test = new Insulation_Test { };

        bool[] testing_Check = { false, false, false, false, false, false, false };
        bool scan_read = false;
        bool isESP_connected = false;
        bool isTemplateSelected = false;

        public Form1()
        {
            InitializeComponent();

            esp32_module = new SerialPort();
            esp32_module.Parity = Parity.None;
            esp32_module.DataBits = 8;
            esp32_module.StopBits = StopBits.One;
            esp32_module.DtrEnable = true;
            esp32_module.RtsEnable = true;
            esp32_module.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);

            receive_present_state = ReportReceive_states.Standby;



            updateESP32_Connection_Status(esp32_module.IsOpen);
            updateStatus_BTN();

            ExportDOCX.Text = "Please Select Template File...";
            ExportDOCX.Enabled = false;
            Templ_from.Text = "-";

            Exp_to.Text = Save_to.SelectedPath;
            SerialMoni.Visible = false;

            //Req_report.Enabled = esp32_module.IsOpen;
        }
        private void hook_Connect_ESP()
        {
            if (esp32_module.IsOpen == false)
            {
                esp32_module.BaudRate = 115200;
                try
                {
                    esp32_module.PortName = COM_Input.Text;
                }
                catch (ArgumentException ex)
                {
                    ESP_Status.Text = ex.Message;
                }


                try
                {
                    esp32_module.Open();
                    updateESP32_Connection_Status(esp32_module.IsOpen);
                }
                catch (FileNotFoundException fe)
                {
                    if (ESP_Status.InvokeRequired)
                    {
                        Action update_connect_status = delegate { updateESP32_Connection_Status(false); };
                        ESP_Status.Invoke(update_connect_status);
                    }
                    else
                    {
                        ESP_Status.Text = fe.Message;

                    }
                }

            }
            else
            {
                esp32_module.Close();
                updateESP32_Connection_Status(esp32_module.IsOpen);
            }
        }
        private void Connect_ESP_BTN_Click(object sender, EventArgs e)
        {
            hook_Connect_ESP();
        }

        private void updateESP32_Connection_Status(bool sts)
        {

            if (ESP_Status.InvokeRequired)
            {
                Action update_connect_status = delegate { updateESP32_Connection_Status(sts); };
                ESP_Status.Invoke(update_connect_status);
            }
            else
            {
                if (sts)
                {
                    ESP_Status.Text = "Connected";
                    Connect_ESP_BTN.Text = "Disconnect to ESP32";
                    isESP_connected = true;
                }
                else
                {
                    ESP_Status.Text = "Not Connected";
                    Connect_ESP_BTN.Text = "Connect to ESP32";
                    isESP_connected = false;
                }

                Enable_All_Test_BTN(sts);
            }


        }

        private void Enable_All_Test_BTN(bool sts)
        {
            if (Test_AB.InvokeRequired)
            {
                Action test_sts_req = delegate { Enable_All_Test_BTN(sts); };
                Test_AB.Invoke(test_sts_req);
            }
            else
            {
                Test_AB.Enabled = sts;
                Test_BC.Enabled = sts;
                Test_CB.Enabled = sts;
                Test_BD.Enabled = sts;
                Test_RCD.Enabled = sts;
                Test_diode.Enabled = sts;
                Test_Insulat.Enabled = sts;
                Test_ALL_BTN.Enabled = sts;
                cancelBTN.Enabled = !sts && isESP_connected;
                update_btn_Text();
            }
        }

        void update_btn_Text()
        {
            if (Test_AB.InvokeRequired)
            {
                Action test_sts_req = delegate { update_btn_Text(); };
                Test_AB.Invoke(test_sts_req);
            }
            else
            {
                switch (receive_present_state)
                {
                    case ReportReceive_states.Standby:
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_Insulat.Text = "Test";
                        break;
                    case ReportReceive_states.Req_AB:
                        Test_AB.Text = "Testing...";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_Insulat.Text = "Test";
                        break;
                    case ReportReceive_states.Req_BC:
                        Test_BC.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_Insulat.Text = "Test";
                        break;
                    case ReportReceive_states.Req_CB:
                        Test_CB.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_Insulat.Text = "Test";
                        break;
                    case ReportReceive_states.Req_BD:
                        Test_BD.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_Insulat.Text = "Test";
                        break;
                    case ReportReceive_states.Req_Diode:
                        Test_diode.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_Insulat.Text = "Test";
                        break;
                    case ReportReceive_states.Req_RCD:
                        Test_RCD.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_Insulat.Text = "Test";
                        break;
                    case ReportReceive_states.Req_Insul:
                        Test_Insulat.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        break;
                }
            }

        }

        private void addTextToSerialMon(String str)
        {
            if (SerialMoni.InvokeRequired)
            {
                Action seri = delegate { addTextToSerialMon(str); };
                SerialMoni.Invoke(seri);
            }
            else
            {
                SerialMoni.AppendText(str);
            }

        }


        private void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            //string indata = sp.ReadExisting();
            string indata = "";
            indata = sp.ReadLine();

            addTextToSerialMon(indata);

            switch (receive_present_state)
            {
                case ReportReceive_states.Standby:
                    break;
                case ReportReceive_states.Req_AB:
                    while (indata.ToCharArray()[0] != '{')
                    {
                        indata = sp.ReadLine();
                    }
                    read_result_A_to_B(JsonSerializer.Deserialize<State_Transition_Test>(indata));
                    testing_Check[0] = true;

                    if (scan_read)
                    {
                        hook_Test_BC();
                        update_btn_Text();
                    }
                    else
                    {
                        receive_present_state = ReportReceive_states.Standby;
                        Enable_All_Test_BTN(true);
                    }

                    break;
                case ReportReceive_states.Req_BC:
                    while (indata.ToCharArray()[0] != '{')
                    {
                        indata = sp.ReadLine();
                    }

                    read_result_B_to_C(JsonSerializer.Deserialize<State_Transition_Test>(indata));
                    testing_Check[1] = true;
                    if (scan_read)
                    {
                        hook_Test_BD();
                        update_btn_Text();
                    }
                    else
                    {
                        receive_present_state = ReportReceive_states.Standby;
                        Enable_All_Test_BTN(true);
                    }

                    break;
                case ReportReceive_states.Req_CB:
                    while (indata.ToCharArray()[0] != '{')
                    {
                        indata = sp.ReadLine();
                    }

                    read_result_C_to_B(JsonSerializer.Deserialize<State_Transition_Test>(indata));
                    testing_Check[2] = true;

                    receive_present_state = ReportReceive_states.Standby;
                    Enable_All_Test_BTN(true);
                    scan_read = false;

                    break;
                case ReportReceive_states.Req_BD://                                                       
                    while (indata.ToCharArray()[0] != '{')
                    {
                        indata = sp.ReadLine();
                    }

                    read_result_B_to_D(JsonSerializer.Deserialize<State_Transition_Test>(indata));
                    testing_Check[3] = true;

                    
                    if (scan_read)
                    {
                        hook_Test_CB();
                        update_btn_Text();
                    }
                    else
                    {
                        receive_present_state = ReportReceive_states.Standby;
                        Enable_All_Test_BTN(true);
                    }

                    break;
                case ReportReceive_states.Req_Diode:
                    while (indata.ToCharArray()[0] != '{')
                    {
                        indata = sp.ReadLine();
                    }

                    read_result_Diode(JsonSerializer.Deserialize<Diode_Test>(indata));
                    testing_Check[4] = true;

                    receive_present_state = ReportReceive_states.Standby;
                    Enable_All_Test_BTN(true);

                    break;
                case ReportReceive_states.Req_RCD:
                    while (indata.ToCharArray()[0] != '{')
                    {
                        indata = sp.ReadLine();
                    }

                    read_result_RCD(JsonSerializer.Deserialize<RCD0>(indata));
                    testing_Check[5] = true;
                    receive_present_state = ReportReceive_states.Standby;
                    Enable_All_Test_BTN(true);
                    break;
                case ReportReceive_states.Req_Insul:
                    while (indata.ToCharArray()[0] != '{')
                    {
                        indata = sp.ReadLine();
                    }

                    read_result_Insulator(JsonSerializer.Deserialize<Insulation_Test>(indata));
                    testing_Check[6] = true;
                    receive_present_state = ReportReceive_states.Standby;
                    Enable_All_Test_BTN(true);
                    break;
            }
            updateStatus_BTN();

        }

        void read_result_A_to_B(State_Transition_Test s_A_B)
        {
            result_State_A_to_B = s_A_B;
            if (AB_PWM_Startup.InvokeRequired)
            {
                Action add_str = delegate
                {
                    read_result_A_to_B(s_A_B);
                };
                AB_PWM_Startup.Invoke(add_str);
            }
            else
            {
                AB_PWM_Startup.Text = result_State_A_to_B.PWM_StartupDelay + " ms";
            }

            AB_PWM_Amp.Text = result_State_A_to_B.PWM_Amplitude + " V";
            AB_PWM_NVE.Text = result_State_A_to_B.PWM_NveAmplitude + " V";
            AB_PWM_Freq.Text = result_State_A_to_B.PWM_Freq + " Hz";
            AB_PWM_Duty.Text = result_State_A_to_B.PWM_DutyCycle + " %";
            AB_PWM_Imax.Text = result_State_A_to_B.PWM_Imax + " A";

            if (result_State_A_to_B.PWM_StartupDelay_Result)
            {
                AB_PWM_Startup_Result.Text = "Pass";
                AB_PWM_Startup_Result.ForeColor = Color.Green;
            }
            else
            {
                AB_PWM_Startup_Result.Text = "Fail";
                AB_PWM_Startup_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_Amplitude_Result)
            {
                AB_PWM_Amp_Result.Text = "Pass";
                AB_PWM_Amp_Result.ForeColor = Color.Green;
            }
            else
            {
                AB_PWM_Amp_Result.Text = "Fail";
                AB_PWM_Amp_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_NveAmplitude_Result)
            {
                AB_PWM_NVE_Result.Text = "Pass";
                AB_PWM_NVE_Result.ForeColor = Color.Green;
            }
            else
            {
                AB_PWM_NVE_Result.Text = "Fail";
                AB_PWM_NVE_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_Freq_Result)
            {
                AB_PWM_Freq_Result.Text = "Pass";
                AB_PWM_Freq_Result.ForeColor = Color.Green;
            }
            else
            {
                AB_PWM_Freq_Result.Text = "Fail";
                AB_PWM_Freq_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_DutyCycle_Result)
            {
                AB_PWM_Duty_Result.Text = "Pass";
                AB_PWM_Duty_Result.ForeColor = Color.Green;
            }
            else
            {
                AB_PWM_Duty_Result.Text = "Fail";
                AB_PWM_Duty_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_Imax_Result)
            {
                AB_PWM_Imax_Result.Text = "Pass";
                AB_PWM_Imax_Result.ForeColor = Color.Green;
            }
            else
            {
                AB_PWM_Imax_Result.Text = "Fail";
                AB_PWM_Imax_Result.ForeColor = Color.Red;
            }

            /*
            if (result_State_A_to_B.Testing_Result)
            {
                AB_check.Text = "Pass";
                AB_check.ForeColor = Color.Green;
            }
            else
            {
                AB_check.Text = "Fail";
                AB_check.ForeColor = Color.Red;
            }
            */
        }
        void read_result_B_to_C(State_Transition_Test s_B_C)
        {
            result_State_B_to_C = s_B_C;
            if (BC_PWM_Startup.InvokeRequired)
            {
                Action add_str = delegate
                {
                    read_result_B_to_C(s_B_C);
                };
                BC_PWM_Startup.Invoke(add_str);
            }
            else
            {
                BC_PWM_Startup.Text = result_State_B_to_C.PWM_StartupDelay + " ms";
            }

            BC_PWM_Amp.Text = result_State_B_to_C.PWM_Amplitude + " V";
            BC_PWM_NVE.Text = result_State_B_to_C.PWM_NveAmplitude + " V";
            BC_PWM_Freq.Text = result_State_B_to_C.PWM_Freq + " Hz";
            BC_PWM_Duty.Text = result_State_B_to_C.PWM_DutyCycle + " %";
            BC_PWM_Imax.Text = result_State_B_to_C.PWM_Imax + " A";
            BC_Voltage.Text = result_State_B_to_C.Voltage + " V";
            /*
            if (result_State_B_to_C.Testing_Result)
            {
                BC_check.Text = "Pass";
                BC_check.ForeColor = Color.Green;
            }
            else
            {
                BC_check.Text = "Fail";
                BC_check.ForeColor = Color.Red;
            }
            */
            BC_PWM_OnDel.Text = result_State_B_to_C.MainsOnDelay + " ms";
            BC_PWM_MainFreq.Text = result_State_B_to_C.MainsFreq + " Hz";
            BC_PP.Text = result_State_B_to_C.PP + " A";

            if (result_State_B_to_C.PWM_StartupDelay_Result)
            {
                BC_PWM_Startup_Result.Text = "Pass";
                BC_PWM_Startup_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PWM_Startup_Result.Text = "Fail";
                BC_PWM_Startup_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_Amplitude_Result)
            {
                BC_PWM_Amp_Result.Text = "Pass";
                BC_PWM_Amp_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PWM_Amp_Result.Text = "Fail";
                BC_PWM_Amp_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_NveAmplitude_Result)
            {
                BC_PWM_NVE_Result.Text = "Pass";
                BC_PWM_NVE_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PWM_NVE_Result.Text = "Fail";
                BC_PWM_NVE_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_Freq_Result)
            {
                BC_PWM_Freq_Result.Text = "Pass";
                BC_PWM_Freq_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PWM_Freq_Result.Text = "Fail";
                BC_PWM_Freq_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_DutyCycle_Result)
            {
                BC_PWM_Duty_Result.Text = "Pass";
                BC_PWM_Duty_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PWM_Duty_Result.Text = "Fail";
                BC_PWM_Duty_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_Imax_Result)
            {
                BC_PWM_Imax_Result.Text = "Pass";
                BC_PWM_Imax_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PWM_Imax_Result.Text = "Fail";
                BC_PWM_Imax_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.Voltage_Result)
            {
                BC_Voltage_Result.Text = "Pass";
                BC_Voltage_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_Voltage_Result.Text = "Fail";
                BC_Voltage_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.MainsOnDelay_Result)
            {
                BC_PWM_OnDel_Result.Text = "Pass";
                BC_PWM_OnDel_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PWM_OnDel_Result.Text = "Fail";
                BC_PWM_OnDel_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.MainsFreq_Result)
            {
                BC_PWM_MainFreq_Result.Text = "Pass";
                BC_PWM_MainFreq_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PWM_MainFreq_Result.Text = "Fail";
                BC_PWM_MainFreq_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PP_Result)
            {
                BC_PP_Result.Text = "Pass";
                BC_PP_Result.ForeColor = Color.Green;
            }
            else
            {
                BC_PP_Result.Text = "Fail";
                BC_PP_Result.ForeColor = Color.Red;
            }

        }
        void read_result_C_to_B(State_Transition_Test s_C_B)
        {
            result_State_C_to_B = s_C_B;
            if (CB_PWM_Startup.InvokeRequired)
            {
                Action add_str = delegate
                {
                    read_result_C_to_B(s_C_B);
                };
                CB_PWM_Startup.Invoke(add_str);
            }
            else
            {
                CB_PWM_Startup.Text = result_State_C_to_B.PWM_StartupDelay + " ms";
            }

            CB_PWM_Amp.Text = result_State_C_to_B.PWM_Amplitude + " V";
            CB_PWM_NVE.Text = result_State_C_to_B.PWM_NveAmplitude + " V";
            CB_PWM_Freq.Text = result_State_C_to_B.PWM_Freq + " Hz";
            CB_PWM_Duty.Text = result_State_C_to_B.PWM_DutyCycle + " %";
            CB_PWM_Imax.Text = result_State_C_to_B.PWM_Imax + " A";
            CB_PWM_OffDel.Text = result_State_C_to_B.MainsOffDelay + " ms";

            if (result_State_C_to_B.PWM_StartupDelay_Result)
            {
                CB_PWM_Startup_Result.Text = "Pass";
                CB_PWM_Startup_Result.ForeColor = Color.Green;
            }
            else
            {
                CB_PWM_Startup_Result.Text = "Fail";
                CB_PWM_Startup_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_Amplitude_Result)
            {
                CB_PWM_Amp_Result.Text = "Pass";
                CB_PWM_Amp_Result.ForeColor = Color.Green;
            }
            else
            {
                CB_PWM_Amp_Result.Text = "Fail";
                CB_PWM_Amp_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_NveAmplitude_Result)
            {
                CB_PWM_NVE_Result.Text = "Pass";
                CB_PWM_NVE_Result.ForeColor = Color.Green;
            }
            else
            {
                CB_PWM_NVE_Result.Text = "Fail";
                CB_PWM_NVE_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_Freq_Result)
            {
                CB_PWM_Freq_Result.Text = "Pass";
                CB_PWM_Freq_Result.ForeColor = Color.Green;
            }
            else
            {
                CB_PWM_Freq_Result.Text = "Fail";
                CB_PWM_Freq_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_DutyCycle_Result)
            {
                CB_PWM_Duty_Result.Text = "Pass";
                CB_PWM_Duty_Result.ForeColor = Color.Green;
            }
            else
            {
                CB_PWM_Duty_Result.Text = "Fail";
                CB_PWM_Duty_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_Imax_Result)
            {
                CB_PWM_Imax_Result.Text = "Pass";
                CB_PWM_Imax_Result.ForeColor = Color.Green;
            }
            else
            {
                CB_PWM_Imax_Result.Text = "Fail";
                CB_PWM_Imax_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.MainsOffDelay_Result)
            {
                CB_PWM_OffDel_Result.Text = "Pass";
                CB_PWM_OffDel_Result.ForeColor = Color.Green;
            }
            else
            {
                CB_PWM_OffDel_Result.Text = "Fail";
                CB_PWM_OffDel_Result.ForeColor = Color.Red;
            }


            // CB_PWM_Startup_Result
            // CB_PWM_Amp_Result
            // CB_PWM_NVE_Result
            // CB_PWM_Freq_Result
            // CB_PWM_Duty_Result
            // CB_PWM_Imax_Result
            // CB_PWM_OffDel_Result

            /*
            if (result_State_C_to_B.Testing_Result)
            {
                CB_check.Text = "Pass";
                CB_check.ForeColor = Color.Green;
            }
            else
            {
                CB_check.Text = "Fail";
                CB_check.ForeColor = Color.Red;
            }*/
        }
        void read_result_B_to_D(State_Transition_Test s_B_D)
        {
            result_State_B_to_D = s_B_D;
            if (BD_PWM_Startup.InvokeRequired)
            {
                Action add_str = delegate
                {
                    read_result_B_to_D(s_B_D);
                };
                BD_PWM_Startup.Invoke(add_str);
            }
            else
            {
                BD_PWM_Startup.Text = result_State_B_to_D.PWM_StartupDelay + " ms";
            }

            BD_PWM_Amp.Text = result_State_B_to_D.PWM_Amplitude + " V";
            BD_PWM_NVE.Text = result_State_B_to_D.PWM_NveAmplitude + " V";
            BD_PWM_Freq.Text = result_State_B_to_D.PWM_Freq + " Hz";
            BD_PWM_Duty.Text = result_State_B_to_D.PWM_DutyCycle + " %";
            BD_PWM_Imax.Text = result_State_B_to_D.PWM_Imax + " A";
            BD_Voltage.Text = result_State_B_to_D.Voltage + " V";
            /*if (result_State_B_to_D.Testing_Result)
            {
                BD_check.Text = "Pass";
                BD_check.ForeColor = Color.Green;
            }
            else
            {
                BD_check.Text = "Fail";
                BD_check.ForeColor = Color.Red;
            }*/
            BD_PWM_OnDel.Text = result_State_B_to_D.MainsOnDelay + " ms";
            BD_PWM_MainFreq.Text = result_State_B_to_D.MainsFreq + " Hz";
            BD_PP.Text = result_State_B_to_D.PP + " A";

            if (result_State_B_to_D.PWM_StartupDelay_Result)
            {
                BD_PWM_Startup_Result.Text = "Pass";
                BD_PWM_Startup_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PWM_Startup_Result.Text = "Fail";
                BD_PWM_Startup_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_Amplitude_Result)
            {
                BD_PWM_Amp_Result.Text = "Pass";
                BD_PWM_Amp_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PWM_Amp_Result.Text = "Fail";
                BD_PWM_Amp_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_NveAmplitude_Result)
            {
                BD_PWM_NVE_Result.Text = "Pass";
                BD_PWM_NVE_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PWM_NVE_Result.Text = "Fail";
                BD_PWM_NVE_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_Freq_Result)
            {
                BD_PWM_Freq_Result.Text = "Pass";
                BD_PWM_Freq_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PWM_Freq_Result.Text = "Fail";
                BD_PWM_Freq_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_DutyCycle_Result)
            {
                BD_PWM_Duty_Result.Text = "Pass";
                BD_PWM_Duty_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PWM_Duty_Result.Text = "Fail";
                BD_PWM_Duty_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_Imax_Result)
            {
                BD_PWM_Imax_Result.Text = "Pass";
                BD_PWM_Imax_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PWM_Imax_Result.Text = "Fail";
                BD_PWM_Imax_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.Voltage_Result)
            {
                BD_Voltage_Result.Text = "Pass";
                BD_Voltage_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_Voltage_Result.Text = "Fail";
                BD_Voltage_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.MainsOnDelay_Result)
            {
                BD_PWM_OnDel_Result.Text = "Pass";
                BD_PWM_OnDel_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PWM_OnDel_Result.Text = "Fail";
                BD_PWM_OnDel_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.MainsFreq_Result)
            {
                BD_PWM_MainFreq_Result.Text = "Pass";
                BD_PWM_MainFreq_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PWM_MainFreq_Result.Text = "Fail";
                BD_PWM_MainFreq_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PP_Result)
            {
                BD_PP_Result.Text = "Pass";
                BD_PP_Result.ForeColor = Color.Green;
            }
            else
            {
                BD_PP_Result.Text = "Fail";
                BD_PP_Result.ForeColor = Color.Red;
            }


            // BD_PWM_Startup_Result
            // BD_PWM_Amp_Result
            // BD_PWM_NVE_Result
            // BD_PWM_Freq_Result
            // BD_PWM_Duty_Result
            // BD_PWM_Imax_Result
            // BD_Voltage_Result
            // BD_PWM_OnDel_Result
            // BD_PWM_MainFreq_Result
            // BD_PP_Result
        }

        void read_result_RCD(RCD0 s_rcd)
        {

            result_rcd_test = s_rcd;

            if (RCD_TripTime.InvokeRequired)
            {
                Action add_str = delegate
                {
                    read_result_RCD(s_rcd);
                };
                RCD_TripTime.Invoke(add_str);
            }
            else
            {
                RCD_TripTime.Text = result_rcd_test.Trip_Time + " ms";
            }
            RCD_Limit.Text = result_rcd_test.Limit + " ms";
            RCD_Current.Text = result_rcd_test.Current + " mA";

            if (result_rcd_test.RCD0_Result)
            {
                RCD_check.Text = "Pass";
                RCD_check.ForeColor = Color.Green;
            }
            else
            {
                RCD_check.Text = "Fail";
                RCD_check.ForeColor = Color.Red;
            }
        }
        void read_result_Diode(Diode_Test s_diode)
        {

            result_diode_test = s_diode;

            if (DiodeShort_Delay.InvokeRequired)
            {
                Action add_str = delegate
                {
                    read_result_Diode(s_diode);
                };
                DiodeShort_Delay.Invoke(add_str);
            }
            else
            {
                DiodeShort_Delay.Text = result_diode_test.Diode_ShortCircuit_MainsOffDelay + " ms";
            }
            PE_Open_Delay.Text = result_diode_test.PE_OpenCircuit_MainsOffDelay + " ms";
            DiodeOpen_Delay.Text = result_diode_test.Diode_OpenCircuit_MainsOffDelay + " ms";

            if (result_diode_test.Diode_ShortCircuit_Result)
            {
                Diode_Short_check.Text = "Pass";
                Diode_Short_check.ForeColor = Color.Green;
            }
            else
            {
                Diode_Short_check.Text = "Fail";
                Diode_Short_check.ForeColor = Color.Red;
            }
            if (result_diode_test.PE_OpenCircuit_Result)
            {
                PE_Open_check.Text = "Pass";
                PE_Open_check.ForeColor = Color.Green;
            }
            else
            {
                PE_Open_check.Text = "Fail";
                PE_Open_check.ForeColor = Color.Red;
            }
            if (result_diode_test.Diode_OpenCircuit_Result)
            {
                DiodeOpen_check.Text = "Pass";
                DiodeOpen_check.ForeColor = Color.Green;
            }
            else
            {
                DiodeOpen_check.Text = "Fail";
                DiodeOpen_check.ForeColor = Color.Red;
            }
        }
        void read_result_Insulator(Insulation_Test s_insu)
        {

            Insulation_Test = s_insu;

            if (Insulator_Limit.InvokeRequired)
            {
                Action add_str = delegate
                {
                    read_result_Insulator(s_insu);
                };
                Insulator_Limit.Invoke(add_str);
            }
            else
            {
                Insulator_Limit.Text = Insulation_Test.N_PE + " Ω";
            }
            Insulator_Result.Text = Insulation_Test.L_PE + " Ω";
            Insulator_Volt.Text = Insulation_Test.Voltage + " V";

            if (Insulation_Test.Insulation_Testing)
            {
                Insu_check.Text = "Pass";
                Insu_check.ForeColor = Color.Green;
            }
            else
            {
                Insu_check.Text = "Fail";
                Insu_check.ForeColor = Color.Red;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            clearS_AB();
            clearS_BC();
            clearS_CB();
            clearS_BD();
            clearDiode();
            clearRCD();
            clearInsu();
        }

        private void updateStatus_BTN()
        {
            //bool en = true;
            bool en = isTemplateSelected;
            /*
            foreach (bool b in testing_Check)
            {
                en &= b;
            }*/


            if (ExportPDF.InvokeRequired)
            {
                Action add_str = delegate
                {
                    updateStatus_BTN();
                };
                ExportPDF.Invoke(add_str);
            }
            else
            {
                // ExportPDF.Enabled = en;
                ExportDOCX.Enabled = en;
            }


        }

        private void Test_AB_Click(object sender, EventArgs e)
        {
            hook_Test_AB();
            Enable_All_Test_BTN(false);
        }

        private void Test_BC_Click(object sender, EventArgs e)
        {
            hook_Test_BC();
            Enable_All_Test_BTN(false);
        }

        private void Test_CB_Click(object sender, EventArgs e)
        {
            hook_Test_CB();
            Enable_All_Test_BTN(false);
        }

        private void Test_BD_Click(object sender, EventArgs e)
        {
            hook_Test_BD();
            Enable_All_Test_BTN(false);
        }

        private void Test_RCD_Click(object sender, EventArgs e)
        {
            hook_Test_RCD();
            Enable_All_Test_BTN(false);
        }

        private void Test_Insulat_Click(object sender, EventArgs e)
        {
            hook_Test_Insulator();
            Enable_All_Test_BTN(false);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            hook_Test_Diode();
            Enable_All_Test_BTN(false);
        }

        private void hook_Test_AB()
        {
            esp32_module.WriteLine("State_A_to_B\n");
            receive_present_state = ReportReceive_states.Req_AB;
        }
        private void hook_Test_BC()
        {
            esp32_module.WriteLine("State_B_to_C\n");
            receive_present_state = ReportReceive_states.Req_BC;
        }
        private void hook_Test_CB()
        {
            esp32_module.WriteLine("State_C_to_B\n");
            receive_present_state = ReportReceive_states.Req_CB;
        }
        private void hook_Test_BD()
        {
            esp32_module.WriteLine("State_B_to_D\n");
            receive_present_state = ReportReceive_states.Req_BD;
        }
        private void hook_Test_RCD()
        {
            esp32_module.WriteLine("RCD0_Test\n");
            receive_present_state = ReportReceive_states.Req_RCD;
        }
        private void hook_Test_Insulator()
        {
            esp32_module.WriteLine("Insulator_Test\n");
            receive_present_state = ReportReceive_states.Req_Insul;
        }
        private void hook_Test_Diode()
        {
            esp32_module.WriteLine("Diode_Test\n");
            receive_present_state = ReportReceive_states.Req_Diode;
        }
        

        private void exp_dir_Click(object sender, EventArgs e)
        {
            //DialogResult result = Save_to.ShowDialog();

            DialogResult result = SaveReportAs.ShowDialog();
            if (SaveReportAs.FileName.IndexOf(".docx") == -1) // add .docx when no .docx found in name
            {
                SaveReportAs.FileName += ".docx";
            }
            if (result == DialogResult.OK)
            {
                // Exp_to.Text = Save_to.SelectedPath;
                Exp_to.Text = SaveReportAs.FileName;
            }
        }

        private void Test_ALL_BTN_Click(object sender, EventArgs e)
        {
            hook_Test_AB();
            scan_read = true;
            // Lock all Test button
            Enable_All_Test_BTN(false);
        }

        private void COM_Input_Enter(object sender, KeyPressEventArgs e)
        {

        }

        private void COM_Input_KeyDown(object sender, KeyEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Enter)
                hook_Connect_ESP();
        }

        private void SerialMoni_TextChanged(object sender, EventArgs e)
        {

        }

        private void cancelBTN_Click(object sender, EventArgs e)
        {
            receive_present_state = ReportReceive_states.Standby;
            Enable_All_Test_BTN(true);
            scan_read = false;
        }

        void clearS_AB()
        {
            if (AB_PWM_Startup.InvokeRequired)
            {
                Action add_str = delegate
                {
                    clearS_AB();
                };
                AB_PWM_Startup.Invoke(add_str);
            }
            else
            {
                AB_PWM_Startup.Text = "-";
            }

            AB_PWM_Amp.Text = "-";
            AB_PWM_NVE.Text = "-";
            AB_PWM_Freq.Text = "-";
            AB_PWM_Duty.Text = "-";
            AB_PWM_Imax.Text = "-";

            AB_PWM_Startup_Result.Text = "(-)";
            AB_PWM_Amp_Result.Text = "(-)";
            AB_PWM_NVE_Result.Text = "(-)";
            AB_PWM_Freq_Result.Text = "(-)";
            AB_PWM_Duty_Result.Text = "(-)";
            AB_PWM_Imax_Result.Text = "(-)";

            AB_PWM_Startup_Result.ForeColor = Color.Black;
            AB_PWM_Amp_Result.ForeColor = Color.Black;
            AB_PWM_NVE_Result.ForeColor = Color.Black;
            AB_PWM_Freq_Result.ForeColor = Color.Black;
            AB_PWM_Duty_Result.ForeColor = Color.Black;
            AB_PWM_Imax_Result.ForeColor = Color.Black;

            AB_check.Text = "(-)";
            AB_check.ForeColor = Color.Black;

        }
        void clearS_BC()
        {
            if (BC_PWM_Startup.InvokeRequired)
            {
                Action add_str = delegate
                {
                    clearS_BC();
                };
                BC_PWM_Startup.Invoke(add_str);
            }
            else
            {
                BC_PWM_Startup.Text = "-";
            }

            BC_PWM_Amp.Text = "-";
            BC_PWM_NVE.Text = "-";
            BC_PWM_Freq.Text = "-";
            BC_PWM_Duty.Text = "-";
            BC_PWM_Imax.Text = "-";
            BC_Voltage.Text = "-";

            BC_check.Text = "(-)";
            BC_check.ForeColor = Color.Black;

            BC_PWM_OnDel.Text = "-";
            BC_PWM_MainFreq.Text = "-";
            BC_PP.Text = "-";

            BC_PWM_Startup_Result.Text = "(-)";
            BC_PWM_Amp_Result.Text = "(-)";
            BC_PWM_NVE_Result.Text = "(-)";
            BC_PWM_Freq_Result.Text = "(-)";
            BC_PWM_Duty_Result.Text = "(-)";
            BC_PWM_Imax_Result.Text = "(-)";
            BC_Voltage_Result.Text = "(-)";
            BC_PWM_OnDel_Result.Text = "(-)";
            BC_PWM_MainFreq_Result.Text = "(-)";
            BC_PP_Result.Text = "(-)";

            BC_PWM_Startup_Result.ForeColor = Color.Black;
            BC_PWM_Amp_Result.ForeColor = Color.Black;
            BC_PWM_NVE_Result.ForeColor = Color.Black;
            BC_PWM_Freq_Result.ForeColor = Color.Black;
            BC_PWM_Duty_Result.ForeColor = Color.Black;
            BC_PWM_Imax_Result.ForeColor = Color.Black;
            BC_Voltage_Result.ForeColor = Color.Black;
            BC_PWM_OnDel_Result.ForeColor = Color.Black;
            BC_PWM_MainFreq_Result.ForeColor = Color.Black;
            BC_PP_Result.ForeColor = Color.Black;
        }
        void clearS_CB()
        {
            if (CB_PWM_Startup.InvokeRequired)
            {
                Action add_str = delegate
                {
                    clearS_CB();
                };
                CB_PWM_Startup.Invoke(add_str);
            }
            else
            {
                CB_PWM_Startup.Text = "-";
            }

            CB_PWM_Amp.Text = "-";
            CB_PWM_NVE.Text = "-";
            CB_PWM_Freq.Text = "-";
            CB_PWM_Duty.Text = "-";
            CB_PWM_Imax.Text = "-";
            CB_PWM_OffDel.Text = "-";

            CB_PWM_Startup_Result.Text = "(-)";
            CB_PWM_Amp_Result.Text = "(-)";
            CB_PWM_NVE_Result.Text = "(-)";
            CB_PWM_Freq_Result.Text = "(-)";
            CB_PWM_Duty_Result.Text = "(-)";
            CB_PWM_Imax_Result.Text = "(-)";
            CB_PWM_OffDel_Result.Text = "(-)";

            CB_PWM_Startup_Result.ForeColor = Color.Black;
            CB_PWM_Amp_Result.ForeColor = Color.Black;
            CB_PWM_NVE_Result.ForeColor = Color.Black;
            CB_PWM_Freq_Result.ForeColor = Color.Black;
            CB_PWM_Duty_Result.ForeColor = Color.Black;
            CB_PWM_Imax_Result.ForeColor = Color.Black;
            CB_PWM_OffDel_Result.ForeColor = Color.Black;

            CB_check.Text = "(-)";
            CB_check.ForeColor = Color.Black;
        }
        void clearS_BD()
        {
            if (BD_PWM_Startup.InvokeRequired)
            {
                Action add_str = delegate
                {
                    clearS_BD();
                };
                BD_PWM_Startup.Invoke(add_str);
            }
            else
            {
                BD_PWM_Startup.Text = "-";
            }

            BD_PWM_Amp.Text = "-";
            BD_PWM_NVE.Text = "-";
            BD_PWM_Freq.Text = "-";
            BD_PWM_Duty.Text = "-";
            BD_PWM_Imax.Text = "-";
            BD_Voltage.Text = "-";

            BD_check.Text = "(-)";
            BD_check.ForeColor = Color.Black;

            BD_PWM_OnDel.Text = "-";
            BD_PWM_MainFreq.Text = "-";
            BD_PP.Text = "-";

            BD_PWM_Startup_Result.Text = "(-)";
            BD_PWM_Amp_Result.Text = "(-)";
            BD_PWM_NVE_Result.Text = "(-)";
            BD_PWM_Freq_Result.Text = "(-)";
            BD_PWM_Duty_Result.Text = "(-)";
            BD_PWM_Imax_Result.Text = "(-)";
            BD_Voltage_Result.Text = "(-)";
            BD_PWM_OnDel_Result.Text = "(-)";
            BD_PWM_MainFreq_Result.Text = "(-)";
            BD_PP_Result.Text = "(-)";

            BD_PWM_Startup_Result.Text = "(-)";
            BD_PWM_Amp_Result.Text = "(-)";
            BD_PWM_NVE_Result.Text = "(-)";
            BD_PWM_Freq_Result.Text = "(-)";
            BD_PWM_Duty_Result.Text = "(-)";
            BD_PWM_Imax_Result.Text = "(-)";
            BD_Voltage_Result.Text = "(-)";
            BD_PWM_OnDel_Result.Text = "(-)";
            BD_PWM_MainFreq_Result.Text = "(-)";
            BD_PP_Result.Text = "(-)";

            BD_PWM_Startup_Result.ForeColor = Color.Black;
            BD_PWM_Amp_Result.ForeColor = Color.Black;
            BD_PWM_NVE_Result.ForeColor = Color.Black;
            BD_PWM_Freq_Result.ForeColor = Color.Black;
            BD_PWM_Duty_Result.ForeColor = Color.Black;
            BD_PWM_Imax_Result.ForeColor = Color.Black;
            BD_Voltage_Result.ForeColor = Color.Black;
            BD_PWM_OnDel_Result.ForeColor = Color.Black;
            BD_PWM_MainFreq_Result.ForeColor = Color.Black;
            BD_PP_Result.ForeColor = Color.Black;
        }
        void clearDiode()
        {
            if (DiodeShort_Delay.InvokeRequired)
            {
                Action add_str = delegate
                {
                    clearDiode();
                };
                DiodeShort_Delay.Invoke(add_str);
            }
            else
            {
                DiodeShort_Delay.Text = "-";
            }
            PE_Open_Delay.Text = "-";
            DiodeOpen_Delay.Text = "-";

            Diode_Short_check.Text = "(-)";
            Diode_Short_check.ForeColor = Color.Black;


            PE_Open_check.Text = "(-)";
            PE_Open_check.ForeColor = Color.Black;


            DiodeOpen_check.Text = "(-)";
            DiodeOpen_check.ForeColor = Color.Black;

        }
        void clearRCD()
        {

            if (RCD_TripTime.InvokeRequired)
            {
                Action add_str = delegate
                {
                    clearRCD();
                };
                RCD_TripTime.Invoke(add_str);
            }
            else
            {
                RCD_TripTime.Text = "-";
            }
            RCD_Limit.Text = "-";
            RCD_Current.Text = "-";

            RCD_check.Text = "(-)";
            RCD_check.ForeColor = Color.Black;

        }
        void clearInsu()
        {
            if (Insulator_Limit.InvokeRequired)
            {
                Action add_str = delegate
                {
                    clearInsu();
                };
                Insulator_Limit.Invoke(add_str);
            }
            else
            {
                Insulator_Limit.Text = "-";
            }
            Insulator_Result.Text = "-";
            Insulator_Volt.Text = "-";

            Insu_check.Text = "(-)";
            Insu_check.ForeColor = Color.Black;

        }

        private void ClrResult_Click(object sender, EventArgs e)
        {
            clearS_AB();
            clearS_BC();
            clearS_CB();
            clearS_BD();
            clearDiode();
            clearRCD();
            clearInsu();
        }

        private void ExportDOCX_Click(object sender, EventArgs e)
        {
            // using(WordDocument doc = new WordDocument())
            // {
            //     doc.EnsureMinimal();
            // 
            //     doc.LastParagraph.AppendText("Hello world");
            // 
            //     doc.Save(Save_to.SelectedPath + "\\" + "EVSE_Test_report.docx");
            //     doc.Close();
            // }
            /*            var value = new Dictionary<string, object>()
                        {
                            ["name"] = "Jack",
                            ["surname"] = "Madison",
                            ["date"] = DateTime.Now.ToString(),
                        };
            */
            SaveReportAs.FileName = "EVSE_Test_report_" + DeviceNo_Test.Text + " " + $"{DateTime.Now:dd-MM-yyyy hh-mm-ss}" + ".docx";
            DialogResult result = SaveReportAs.ShowDialog();
            string saveto_dir = "";
            // If selected save dir then savefile otherwise do nothing
            if (result == DialogResult.OK)
            {
                /*
                MiniWordColorText ab_testing_result;
                MiniWordColorText bc_testing_result;
                MiniWordColorText cb_testing_result;
                MiniWordColorText bd_testing_result;

                if (result_State_A_to_B.Testing_Result) { ab_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { ab_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                if (result_State_B_to_C.Testing_Result) { bc_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { bc_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                if (result_State_C_to_B.Testing_Result) { cb_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { cb_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                if (result_State_B_to_D.Testing_Result) { bd_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { bd_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                */



                MiniWordColorText AB_PWM_Startup_Text_Word;
                MiniWordColorText AB_PWM_Amp_Text_Word;
                MiniWordColorText AB_PWM_NVE_Text_Word;
                MiniWordColorText AB_PWM_Freq_Text_Word;
                MiniWordColorText AB_PWM_Duty_Text_Word;
                MiniWordColorText AB_PWM_Imax_Text_Word;
                MiniWordColorText BC_PWM_Startup_Text_Word;
                MiniWordColorText BC_PWM_Amp_Text_Word;
                MiniWordColorText BC_PWM_NVE_Text_Word;
                MiniWordColorText BC_PWM_Freq_Text_Word;
                MiniWordColorText BC_PWM_Duty_Text_Word;
                MiniWordColorText BC_PWM_Imax_Text_Word;
                MiniWordColorText BC_Voltage_Text_Word;
                MiniWordColorText BC_PWM_OnDel_Text_Word;
                MiniWordColorText BC_PWM_MainFreq_Text_Word;
                MiniWordColorText BC_PP_Text_Word;
                MiniWordColorText CB_PWM_Startup_Text_Word;
                MiniWordColorText CB_PWM_Amp_Text_Word;
                MiniWordColorText CB_PWM_NVE_Text_Word;
                MiniWordColorText CB_PWM_Freq_Text_Word;
                MiniWordColorText CB_PWM_Duty_Text_Word;
                MiniWordColorText CB_PWM_Imax_Text_Word;
                MiniWordColorText CB_PWM_OffDel_Text_Word;
                MiniWordColorText BD_PWM_Startup_Text_Word;
                MiniWordColorText BD_PWM_Amp_Text_Word;
                MiniWordColorText BD_PWM_NVE_Text_Word;
                MiniWordColorText BD_PWM_Freq_Text_Word;
                MiniWordColorText BD_PWM_Duty_Text_Word;
                MiniWordColorText BD_PWM_Imax_Text_Word;
                MiniWordColorText BD_Voltage_Text_Word;
                MiniWordColorText BD_PWM_OnDel_Text_Word;
                MiniWordColorText BD_PWM_MainFreq_Text_Word;
                MiniWordColorText BD_PP_Text_Word;

                //                if (result_State_B_to_D.Testing_Result) { bd_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                //                else { bd_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                if (result_State_A_to_B.PWM_StartupDelay_Result) { AB_PWM_Startup_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { AB_PWM_Startup_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_A_to_B.PWM_Amplitude_Result) { AB_PWM_Amp_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { AB_PWM_Amp_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_A_to_B.PWM_NveAmplitude_Result) { AB_PWM_NVE_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { AB_PWM_NVE_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_A_to_B.PWM_Freq_Result) { AB_PWM_Freq_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { AB_PWM_Freq_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_A_to_B.PWM_DutyCycle_Result) { AB_PWM_Duty_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { AB_PWM_Duty_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_A_to_B.PWM_Imax_Result) { AB_PWM_Imax_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { AB_PWM_Imax_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.PWM_StartupDelay_Result) { BC_PWM_Startup_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PWM_Startup_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.PWM_Amplitude_Result) { BC_PWM_Amp_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PWM_Amp_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.PWM_NveAmplitude_Result) { BC_PWM_NVE_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PWM_NVE_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.PWM_Freq_Result) { BC_PWM_Freq_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PWM_Freq_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.PWM_DutyCycle_Result) { BC_PWM_Duty_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PWM_Duty_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.PWM_Imax_Result) { BC_PWM_Imax_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PWM_Imax_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.Voltage_Result) { BC_Voltage_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_Voltage_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.MainsOnDelay_Result) { BC_PWM_OnDel_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PWM_OnDel_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.MainsFreq_Result) { BC_PWM_MainFreq_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PWM_MainFreq_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_C.PP_Result) { BC_PP_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BC_PP_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_C_to_B.PWM_StartupDelay_Result) { CB_PWM_Startup_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { CB_PWM_Startup_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_C_to_B.PWM_Amplitude_Result) { CB_PWM_Amp_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { CB_PWM_Amp_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_C_to_B.PWM_NveAmplitude_Result) { CB_PWM_NVE_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { CB_PWM_NVE_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_C_to_B.PWM_Freq_Result) { CB_PWM_Freq_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { CB_PWM_Freq_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_C_to_B.PWM_DutyCycle_Result) { CB_PWM_Duty_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { CB_PWM_Duty_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_C_to_B.PWM_Imax_Result) { CB_PWM_Imax_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { CB_PWM_Imax_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_C_to_B.MainsOffDelay_Result) { CB_PWM_OffDel_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { CB_PWM_OffDel_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.PWM_StartupDelay_Result) { BD_PWM_Startup_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PWM_Startup_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.PWM_Amplitude_Result) { BD_PWM_Amp_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PWM_Amp_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.PWM_NveAmplitude_Result) { BD_PWM_NVE_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PWM_NVE_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.PWM_Freq_Result) { BD_PWM_Freq_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PWM_Freq_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.PWM_DutyCycle_Result) { BD_PWM_Duty_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PWM_Duty_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.PWM_Imax_Result) { BD_PWM_Imax_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PWM_Imax_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.Voltage_Result) { BD_Voltage_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_Voltage_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.MainsOnDelay_Result) { BD_PWM_OnDel_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PWM_OnDel_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.MainsFreq_Result) { BD_PWM_MainFreq_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PWM_MainFreq_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_State_B_to_D.PP_Result) { BD_PP_Text_Word = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { BD_PP_Text_Word = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                //MiniWordColorText 

                MiniWordColorText diode_sh_result;
                MiniWordColorText pe_op_result;
                MiniWordColorText diode_op_result;

                if (result_diode_test.Diode_ShortCircuit_Result) { diode_sh_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { diode_sh_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_diode_test.PE_OpenCircuit_Result) { pe_op_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { pe_op_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_diode_test.Diode_OpenCircuit_Result) { diode_op_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { diode_op_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                MiniWordColorText rcd_testing_result;
                if (result_rcd_test.RCD0_Result) { rcd_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { rcd_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                MiniWordColorText insulation_testing_result;
                if (Insulation_Test.Insulation_Testing) { insulation_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; }
                else { insulation_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }



                var value = new Dictionary<string, object>()
                {
                    // State A to B testing Result
                    //["AB_Result"] = ab_testing_result,
                    ["AB_StDel"] = result_State_A_to_B.PWM_StartupDelay + " ms",
                    ["AB_Amp"] = result_State_A_to_B.PWM_Amplitude + " V",
                    ["AB_NveAmp"] = result_State_A_to_B.PWM_NveAmplitude + " V",
                    ["AB_Freq"] = result_State_A_to_B.PWM_Freq + " Hz",
                    ["AB_Duty"] = result_State_A_to_B.PWM_DutyCycle + " %",
                    ["AB_Imx"] = result_State_A_to_B.PWM_Imax + " A",

                    ["AB_StDel_Result"] = AB_PWM_Startup_Text_Word,
                    ["AB_Amp_Result"] = AB_PWM_Amp_Text_Word,
                    ["AB_NveAmp_Result"] = AB_PWM_NVE_Text_Word,
                    ["AB_Freq_Result"] = AB_PWM_Freq_Text_Word,
                    ["AB_Duty_Result"] = AB_PWM_Duty_Text_Word,
                    ["AB_Imx_Result"] = AB_PWM_Imax_Text_Word,

                    // State B to C testing Result
                    //["BC_Result"] = bc_testing_result,
                    ["BC_StDel"] = result_State_B_to_C.PWM_StartupDelay + " ms",
                    ["BC_Amp"] = result_State_B_to_C.PWM_Amplitude + " V",
                    ["BC_NveAmp"] = result_State_B_to_C.PWM_NveAmplitude + " V",
                    ["BC_Freq"] = result_State_B_to_C.PWM_Freq + " Hz",
                    ["BC_Duty"] = result_State_B_to_C.PWM_DutyCycle + " %",
                    ["BC_Imx"] = result_State_B_to_C.PWM_Imax + " A",
                    ["BC_MainOnDel"] = result_State_B_to_C.MainsOnDelay + " ms",
                    ["BC_MainFreq"] = result_State_B_to_C.MainsFreq + " Hz",
                    ["BC_Volt"] = result_State_B_to_C.Voltage + " V",
                    ["BC_PP"] = result_State_B_to_C.PP + " A",

                    ["BC_StDel_Result"] = BC_PWM_Startup_Text_Word,
                    ["BC_Amp_Result"] = BC_PWM_Amp_Text_Word,
                    ["BC_NveAmp_Result"] = BC_PWM_NVE_Text_Word,
                    ["BC_Freq_Result"] = BC_PWM_Freq_Text_Word,
                    ["BC_Duty_Result"] = BC_PWM_Duty_Text_Word,
                    ["BC_Imx_Result"] = BC_PWM_Imax_Text_Word,
                    ["BC_MainOnDel_Result"] = BC_Voltage_Text_Word,
                    ["BC_MainFreq_Result"] = BC_PWM_OnDel_Text_Word,
                    ["BC_Volt_Result"] = BC_PWM_MainFreq_Text_Word,
                    ["BC_PP_Result"] = BC_PP_Text_Word,

                    // State C to B testing Result
                    //["CB_Result"] = cb_testing_result,
                    ["CB_StDel"] = result_State_C_to_B.PWM_StartupDelay + " ms",
                    ["CB_Amp"] = result_State_C_to_B.PWM_Amplitude + " V",
                    ["CB_NveAmp"] = result_State_C_to_B.PWM_NveAmplitude + " V",
                    ["CB_Freq"] = result_State_C_to_B.PWM_Freq + " Hz",
                    ["CB_Duty"] = result_State_C_to_B.PWM_DutyCycle + " %",
                    ["CB_Imx"] = result_State_C_to_B.PWM_Imax + " A",
                    ["CB_MainOffDel"] = result_State_C_to_B.MainsOffDelay + " ms",

                    ["CB_StDel_Result"] = CB_PWM_Startup_Text_Word,
                    ["CB_Amp_Result"] = CB_PWM_Amp_Text_Word,
                    ["CB_NveAmp_Result"] = CB_PWM_NVE_Text_Word,
                    ["CB_Freq_Result"] = CB_PWM_Freq_Text_Word,
                    ["CB_Duty_Result"] = CB_PWM_Duty_Text_Word,
                    ["CB_Imx_Result"] = CB_PWM_Imax_Text_Word,
                    ["CB_MainOffDel_Result"] = CB_PWM_OffDel_Text_Word,

                    // State B to D testing Result
                    //["BD_Result"] = bd_testing_result,
                    ["BD_StDel"] = result_State_B_to_D.PWM_StartupDelay + " ms",
                    ["BD_Amp"] = result_State_B_to_D.PWM_Amplitude + " V",
                    ["BD_NveAmp"] = result_State_B_to_D.PWM_NveAmplitude + " V",
                    ["BD_Freq"] = result_State_B_to_D.PWM_Freq + " Hz",
                    ["BD_Duty"] = result_State_B_to_D.PWM_DutyCycle + " %",
                    ["BD_Imx"] = result_State_B_to_D.PWM_Imax + " A",
                    ["BD_MainOnDel"] = result_State_B_to_D.MainsOnDelay + " ms",
                    ["BD_MainFreq"] = result_State_B_to_D.MainsFreq + " Hz",
                    ["BD_Volt"] = result_State_B_to_D.Voltage + " V",
                    ["BD_PP"] = result_State_B_to_D.PP + " A",

                    ["BD_StDel_Result"] = BD_PWM_Startup_Text_Word,
                    ["BD_Amp_Result"] = BD_PWM_Amp_Text_Word,
                    ["BD_NveAmp_Result"] = BD_PWM_NVE_Text_Word,
                    ["BD_Freq_Result"] = BD_PWM_Freq_Text_Word,
                    ["BD_Duty_Result"] = BD_PWM_Duty_Text_Word,
                    ["BD_Imx_Result"] = BD_PWM_Imax_Text_Word,
                    ["BD_MainOnDel_Result"] = BD_Voltage_Text_Word,
                    ["BD_MainFreq_Result"] = BD_PWM_OnDel_Text_Word,
                    ["BD_Volt_Result"] = BD_PWM_MainFreq_Text_Word,
                    ["BD_PP_Result"] = BD_PP_Text_Word,

                    // Diode Testing Result
                    ["D_Sh_result"] = diode_sh_result,
                    ["D_Sh_OffDel"] = result_diode_test.Diode_ShortCircuit_MainsOffDelay + " ms",
                    ["PE_Op_result"] = pe_op_result,
                    ["PE_Op_OffDel"] = result_diode_test.PE_OpenCircuit_MainsOffDelay + " ms",
                    ["D_Op_result"] = diode_op_result,
                    ["D_Op_OffDel"] = result_diode_test.Diode_OpenCircuit_MainsOffDelay + " ms",

                    // RCD Testing Result
                    ["RCD_result"] = rcd_testing_result,
                    ["RCD_TripTime"] = result_rcd_test.Trip_Time + " ms",
                    ["RCD_Lim"] = result_rcd_test.Limit + " ms",
                    ["RCD_Current"] = result_rcd_test.Current + " mA",

                    // Insulation Testing Result
                    ["Insu_result"] = insulation_testing_result,
                    ["Insu_L_PE"] = Insulation_Test.L_PE + " Ω",
                    ["Insu_N_PE"] = Insulation_Test.N_PE + " Ω",
                    ["Insu_Volt"] = Insulation_Test.Voltage + " V",

                    // Testing Info
                    ["Date_testing"] = $"{DateTime.Now:dd/MM/yyyy}",//DateTime.Now;
                    ["Substation_Name"] = Substation_Test.Text,
                    ["Device_Number"] = DeviceNo_Test.Text,
                    ["Customer_Name"] = Customer_name.Text,
                    ["Tester_Name"] = Tester_name.Text,

                    // Charger Info
                    ["mfr"] = Chg_MFR.Text,
                    ["type"] = Chg_Typ.Text,
                    ["year_mfr"] = Chg_YOMFR.Text,
                    ["serial_no"] = Chg_SerNo.Text,
                    ["Irated"] = Chg_Irated.Text,
                    ["VratedA"] = Chg_VaRated.Text,
                    ["VratedB"] = Chg_VbRated.Text,
                    ["VratedC"] = Chg_VcRated.Text,
                    ["Frated"] = Chg_FreqRated.Text
                };
                //MiniWord.SaveAsByTemplate(Save_to.SelectedPath + "\\" + "EVSE_Test_report.docx", WordTemplate.FileName, value);
                //Exp_to.Text = Save_to.SelectedPath;


                if (SaveReportAs.FileName.IndexOf(".docx") == -1) // add .docx when no .docx found in name
                {
                    SaveReportAs.FileName += ".docx";
                }
                saveto_dir = SaveReportAs.FileName;

                MiniWord.SaveAsByTemplate(saveto_dir, WordTemplate.FileName, value);
                Exp_to.Text = saveto_dir;
            }

        }

        private void SelTelp_Click(object sender, EventArgs e)
        {
            DialogResult result = WordTemplate.ShowDialog();
            if (result == DialogResult.OK)
            {
                ExportDOCX.Text = "Export Report DOCX";
                ExportDOCX.Enabled = true;
                //Exp_to.Text = Save_to.SelectedPath;
                Templ_from.Text = WordTemplate.FileName;
            }
        }

        private void Save_to_HelpRequest(object sender, EventArgs e)
        {

        }

        private void Select_File_word(object sender, CancelEventArgs e)
        {
            isTemplateSelected = true;
            ExportDOCX.Text = "Export Report DOCX";
            ExportDOCX.Enabled = true;
            //Exp_to.Text = Save_to.SelectedPath;
            Templ_from.Text = WordTemplate.FileName;
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }



        /*
private void Req_report_Click(object sender, EventArgs e)
{
var Package_request = new Request_Testing_Result
{
Package_req = "State_B_to_C"
};
}*/
        private void ExportPDF_Click(object sender, EventArgs e)
        {
            //ExportPDF.Text = Save_to.SelectedPath;
            String html_to_export = "";

            html_to_export += "<p><img style=\"float: right;\" src=\"https://www.eng.kmutnb.ac.th/wp-content/uploads/2019/08/LOGO-KMUTNB--300x300.png\" alt=\"\" width=\"200\" height=\"200\" /></p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">&nbsp;</p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">King Mongkut's University of Technology North Bangkok</p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">1518 Pracharat 1 Road, Wongsawang, Bangsue, Bangkok 10800</p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">Tel. 02-555-2000 Ext. 8518-8520</p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">Email: <a href=\"mailto:teratam.b@eng.kmutnb.ac.th\">teratam.b@eng.kmutnb.ac.th</a></p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<hr style=\"padding-left: 60px;\" /><hr style=\"padding-left: 40px;\" />";
            html_to_export += "<h2 style=\"padding-left: 60px; text-align: justify;\"><strong>Power System Laboratory</strong></h2>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">Department of Engineering and Computer Engineering</p>";
            html_to_export += "<hr style=\"padding-left: 60px;\" />";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">&nbsp;</p>";
            html_to_export += "<h1 style=\"padding-left: 60px; text-align: justify;\">รายงานผลการทดสอบเครื่องชาร์จรถยนต์ไฟฟ้าแบบกระแสสลับ</h1>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">&nbsp;</p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">เสนอ</p>";
            html_to_export += "<h3 style=\"padding-left: 60px; text-align: justify;\">บริษัท IBS Corporation จำกัด</h4>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">&nbsp;</p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">โดย</p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">ภาควิชาวิศวกรรมไฟฟ้าและคอมพิวเตอร์ คณะวิศวกรรมศาสตร์</p>";
            html_to_export += "<p style=\"padding-left: 60px; text-align: justify;\">มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ</p>";

            html_to_export += "<p>&#13;</p>";

            html_to_export += "<table style=\"height: 181px; width: 100%; border-collapse: collapse; margin-left: auto; margin-right: auto;\" border=\"1\">";
            html_to_export += "<tbody>";
            html_to_export += "<tr style=\"height: 73px;\">";
            html_to_export += "<td style=\"width: 50%; text-align: center; height: 73px;\" colspan=\"2\">";
            html_to_export += "<h1>State A to B : ";
            /*
            if (result_State_A_to_B.Testing_Result)
            {
                html_to_export += "Pass";
            }
            else
            {
                html_to_export += "Fail";
            }
            */
            html_to_export += "</h1>";
            html_to_export += "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Startup Delay</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_A_to_B.PWM_StartupDelay + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Amplitude</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_A_to_B.PWM_Amplitude + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Nve Amplitude</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_A_to_B.PWM_NveAmplitude + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Freq</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_A_to_B.PWM_Freq + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Duty Cycle</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_A_to_B.PWM_DutyCycle + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Imax</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_A_to_B.PWM_Imax + "</td>";
            html_to_export += "</tr>";
            html_to_export += "</tbody>";
            html_to_export += "</table>";

            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";

            html_to_export += "<table style=\"height: 181px; width: 100%; border-collapse: collapse; margin-left: auto; margin-right: auto;\" border=\"1\">";
            html_to_export += "<tbody>";
            html_to_export += "<tr style=\"height: 73px;\">";
            html_to_export += "<td style=\"width: 50%; text-align: center; height: 73px;\" colspan=\"2\">";
            html_to_export += "<h1>State B to C : ";
            /*
            if (result_State_B_to_C.Testing_Result)
            {
                html_to_export += "Pass";
            }
            else
            {
                html_to_export += "Fail";
            }*/
            html_to_export += "</h1>";
            html_to_export += "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Startup Delay</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_B_to_C.PWM_StartupDelay + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Amplitude</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_B_to_C.PWM_Amplitude + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Nve Amplitude</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_B_to_C.PWM_NveAmplitude + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Freq</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_B_to_C.PWM_Freq + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Duty Cycle</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_B_to_C.PWM_DutyCycle + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Imax</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_B_to_C.PWM_Imax + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">Mains On Delay</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_B_to_C.MainsOnDelay + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">Mains Freq</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_B_to_C.MainsFreq + "</td>";
            html_to_export += "</tr>";
            html_to_export += "</tbody>";
            html_to_export += "</table>";

            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";

            html_to_export += "<table style=\"height: 181px; width: 100%; border-collapse: collapse; margin-left: auto; margin-right: auto;\" border=\"1\">";
            html_to_export += "<tbody>";
            html_to_export += "<tr style=\"height: 73px;\">";
            html_to_export += "<td style=\"width: 50%; text-align: center; height: 73px;\" colspan=\"2\">";
            html_to_export += "<h1>State C to B : ";
            /*
            if (result_State_C_to_B.Testing_Result)
            {
                html_to_export += "Pass";
            }
            else
            {
                html_to_export += "Fail";
            }
            */
            html_to_export += "</h1>";
            html_to_export += "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Startup Delay</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_C_to_B.PWM_StartupDelay + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Amplitude</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_C_to_B.PWM_Amplitude + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Nve Amplitude</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_C_to_B.PWM_NveAmplitude + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Freq</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_C_to_B.PWM_Freq + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Duty Cycle</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_C_to_B.PWM_DutyCycle + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">PWM Imax</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_State_C_to_B.PWM_Imax + "</td>";
            html_to_export += "</tr>";
            html_to_export += "</tbody>";
            html_to_export += "</table>";

            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";

            html_to_export += "<table style=\"border-collapse: collapse; width: 100%; height: 273px;\" border=\"1\">";
            html_to_export += "<tbody>";
            html_to_export += "<tr style=\"height: 73px;\">";
            html_to_export += "<td style=\"width: 50%; text-align: center; height: 73px;\" colspan=\"2\">";
            html_to_export += "<h1>Diode Short Circuit : ";
            if (result_diode_test.Diode_ShortCircuit_Result)
            {
                html_to_export += "Pass";
            }
            else
            {
                html_to_export += "Fail";
            }
            html_to_export += "</h1>";
            html_to_export += "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; height: 18px; padding-left: 40px;\">Mains Off Delay</td>";
            html_to_export += "<td style=\"width: 50%; height: 18px; padding-left: 40px;\">" + result_diode_test.Diode_ShortCircuit_MainsOffDelay + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 73px;\">";
            html_to_export += "<td style=\"width: 50%; text-align: center; height: 73px;\" colspan=\"2\">";
            html_to_export += "<h1>Diode Short Circuit Ground : ";
            if (result_diode_test.PE_OpenCircuit_Result)
            {
                html_to_export += "Pass";
            }
            else
            {
                html_to_export += "Fail";
            }
            html_to_export += "</h1>";
            html_to_export += "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; height: 18px; padding-left: 40px;\">Mains Off Delay</td>";
            html_to_export += "<td style=\"width: 50%; height: 18px; padding-left: 40px;\">" + result_diode_test.PE_OpenCircuit_MainsOffDelay + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 73px;\">";
            html_to_export += "<td style=\"width: 50%; text-align: center; height: 73px;\" colspan=\"2\">";
            html_to_export += "<h1>Diode Open Circuit : ";
            if (result_diode_test.Diode_OpenCircuit_Result)
            {
                html_to_export += "Pass";
            }
            else
            {
                html_to_export += "Fail";
            }
            html_to_export += "</h1>";
            html_to_export += "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; height: 18px; padding-left: 40px;\">Mains Off Delay</td>";
            html_to_export += "<td style=\"width: 50%; height: 18px; padding-left: 40px;\">" + result_diode_test.Diode_OpenCircuit_MainsOffDelay + "</td>";
            html_to_export += "</tr>";
            html_to_export += "</tbody>";
            html_to_export += "</table>";

            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";

            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";

            html_to_export += "<p>&nbsp;</p>";

            html_to_export += "<table style=\"height: 181px; width: 100%; border-collapse: collapse; margin-left: auto; margin-right: auto;\" border=\"1\">";
            html_to_export += "<tbody>";
            html_to_export += "<tr style=\"height: 73px;\">";
            html_to_export += "<td style=\"width: 50%; text-align: center; height: 73px;\" colspan=\"2\">";
            html_to_export += "<h1>Rcd0 : ";
            if (result_rcd_test.RCD0_Result)
            {
                html_to_export += "Pass";
            }
            else
            {
                html_to_export += "Fail";
            }
            html_to_export += "</h1>";
            html_to_export += "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">Trip Time</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_rcd_test.Trip_Time + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">Limit</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_rcd_test.Limit + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">Current</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + result_rcd_test.Current + "</td>";
            html_to_export += "</tr>";
            html_to_export += "</tbody>";
            html_to_export += "</table>";

            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";
            html_to_export += "<p>&nbsp;</p>";

            html_to_export += "<table style=\"height: 181px; width: 100%; border-collapse: collapse; margin-left: auto; margin-right: auto;\" border=\"1\">";
            html_to_export += "<tbody>";
            html_to_export += "<tr style=\"height: 73px;\">";
            html_to_export += "<td style=\"width: 50%; text-align: center; height: 73px;\" colspan=\"2\">";
            html_to_export += "<h1>Insulation Test : ";
            if (Insulation_Test.Insulation_Testing)
            {
                html_to_export += "Pass";
            }
            else
            {
                html_to_export += "Fail";
            }
            html_to_export += "</h1>";
            html_to_export += "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">L-PE</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + Insulation_Test.L_PE + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">N-PE</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + Insulation_Test.N_PE + "</td>";
            html_to_export += "</tr>";
            html_to_export += "<tr style=\"height: 18px;\">";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">Voltage</td>";
            html_to_export += "<td style=\"width: 50%; padding-left: 40px; height: 18px;\">" + Insulation_Test.Voltage + "</td>";
            html_to_export += "</tr>";
            html_to_export += "</tbody>";
            html_to_export += "</table>";

            /*
            <table style="border-collapse: collapse; width: 30%; height: 100px;" border="1">
            <tbody>
            <tr style="height: 19px;">
            <td style="width: 50%; text-align: center; height: 19px;">Value</td>
            <td style="width: 50%; text-align: center; height: 19px;">Data</td>
            </tr>
            <tr style="height: 31px;">
            <td style="width: 50%; text-align: center; height: 31px;">Val 1</td>
            <td style="width: 50%; text-align: center; height: 31px;">&nbsp;</td>
            </tr>
            <tr style="height: 31px;">
            <td style="width: 50%; text-align: center; height: 31px;">Val 2</td>
            <td style="width: 50%; text-align: center; height: 31px;">&nbsp;</td>
            </tr>
            <tr style="height: 31px;">
            <td style="width: 50%; text-align: center; height: 31px;">Val 3</td>
            <td style="width: 50%; text-align: center; height: 31px;">&nbsp;</td>
            </tr>
            </tbody>
            </table>
             
             */

            //var pdf_file = new ChromePdfRenderer();
            // instantiate a html to pdf converter object
            HtmlToPdf pdf_file = new HtmlToPdf();
            pdf_file.Options.PdfPageSize = PdfPageSize.A4;
            pdf_file.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            pdf_file.Options.MarginBottom = 14;
            pdf_file.Options.MarginTop = 12;
            pdf_file.Options.MarginLeft = 10;
            pdf_file.Options.MarginRight = 10;


            //= pdf_file.RenderHtmlAsPdf(html_to_export);
            PdfDocument doc = pdf_file.ConvertHtmlString(html_to_export);



            doc.Save(Save_to.SelectedPath + "\\" + "EVSE_Test_report.pdf");
            doc.Close();
        }
    }

    public class Request_Testing_Result
    {
        public string? Package_req { get; set; }
    }

    public class State_Transition_Test
    {
        public string? State_To_Test { get; set; }
        // public bool Testing_Result { get; set; }
        public string? PWM_StartupDelay { get; set; }
        public string? PWM_Amplitude { get; set; }
        public string? PWM_NveAmplitude { get; set; }
        public string? PWM_Freq { get; set; }
        public string? PWM_DutyCycle { get; set; }
        public string? PWM_Imax { get; set; }
        public string? MainsOnDelay { get; set; }
        public string? MainsOffDelay { get; set; }
        public string? MainsFreq { get; set; }
        public string? Voltage { get; set; }
        public string? PP { get; set; }
        public bool PWM_StartupDelay_Result { get; set; }
        public bool PWM_Amplitude_Result { get; set; }
        public bool PWM_NveAmplitude_Result { get; set; }
        public bool PWM_Freq_Result { get; set; }
        public bool PWM_DutyCycle_Result { get; set; }
        public bool PWM_Imax_Result { get; set; }
        public bool MainsOnDelay_Result { get; set; }
        public bool MainsOffDelay_Result { get; set; }
        public bool MainsFreq_Result { get; set; }
        public bool Voltage_Result { get; set; }
        public bool PP_Result { get; set; }
    }

    public class Diode_Test
    {
        public bool Diode_ShortCircuit_Result { get; set; }
        public bool PE_OpenCircuit_Result { get; set; }
        public bool Diode_OpenCircuit_Result { get; set; }
        public string? Diode_ShortCircuit_MainsOffDelay { get; set; }
        public string? PE_OpenCircuit_MainsOffDelay { get; set; }
        public string? Diode_OpenCircuit_MainsOffDelay { get; set; }
    }
    public class RCD0
    {
        public bool RCD0_Result { get; set; }
        public string? Trip_Time { get; set; }
        public string? Limit { get; set; }
        public string? Current { get; set; }
    }

    public class Insulation_Test
    {
        public bool Insulation_Testing { get; set; }
        public string? L_PE { get; set; }
        public string? N_PE { get; set; }
        public string? Voltage { get; set; }
    }

    enum ReportReceive_states
    {
        Standby,
        Req_AB,
        Req_BC,
        Req_CB,
        Req_BD,
        Req_RCD,
        Req_Insul,
        Req_Diode
    }


}