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

using InTheHand.Net.Sockets;
using InTheHand.Net.Bluetooth;
using InTheHand.Net;
using System.Text;

using System.Numerics;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace EV_Testing_report_DEMO
{

    public partial class Form1 : Form {
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

        int num = 0;

        bool[] testing_Check = { false, false, false, false, false, false, false };
        bool scan_read = false;
        bool isESP_connected = false;
        bool isTemplateSelected = false;

        EVSE_Tester_CommunicationMode commu_mode = EVSE_Tester_CommunicationMode.None;
        Diode_Test_ENUM diode_TestCMD = Diode_Test_ENUM.notTesting;

        EVSE_Manual_State manual_State = EVSE_Manual_State.EV_State_A;

        // Create new bluetooth device object
        BluetoothClient client;
        BluetoothDeviceInfo[] devices;
        BluetoothDeviceInfo Selected_Device;
        Boolean bluetooth_connecting = false;

        Stream bluetoothStream;
        byte[] receiveBuffer = new byte[512];
        float[] cp_sample_AB;
        float[] cp_sample_BC;
        float[] cp_sample_CD;

        public Form1() {
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

            client = new BluetoothClient();
            scanBluetooth();

            //Req_report.Enabled = esp32_module.IsOpen;
        }
        private void hook_Connect_ESP() {
            if (esp32_module.IsOpen == false) {
                esp32_module.BaudRate = 115200;
                try {
                    esp32_module.PortName = COM_Input.Text;
                } catch (ArgumentException ex) {
                    ESP_Status.Text = ex.Message;
                }


                try {
                    esp32_module.Open();
                    updateESP32_Connection_Status(esp32_module.IsOpen);
                } catch (FileNotFoundException fe) {
                    if (ESP_Status.InvokeRequired) {
                        Action update_connect_status = delegate { updateESP32_Connection_Status(false); };
                        ESP_Status.Invoke(update_connect_status);
                    } else {
                        ESP_Status.Text = fe.Message;

                    }
                }

                commu_mode = EVSE_Tester_CommunicationMode.SerialPort;

            } else {
                esp32_module.Close();
                updateESP32_Connection_Status(esp32_module.IsOpen);
                commu_mode = EVSE_Tester_CommunicationMode.None;
            }
        }
        private void Connect_ESP_BTN_Click(object sender, EventArgs e) {
            hook_Connect_ESP();
        }

        private void updateESP32_Connection_Status(bool sts) {

            if (ESP_Status.InvokeRequired) {
                Action update_connect_status = delegate { updateESP32_Connection_Status(sts); };
                ESP_Status.Invoke(update_connect_status);
            } else {
                if (sts) {
                    ESP_Status.Text = "Connected";
                    Connect_ESP_BTN.Text = "Disconnect to ESP32";
                    isESP_connected = true;
                } else {
                    ESP_Status.Text = "Not Connected";
                    Connect_ESP_BTN.Text = "Connect to ESP32";
                    isESP_connected = false;
                }

                Enable_All_Test_BTN(sts);
            }


        }

        private void Enable_All_Test_BTN(bool sts) {
            if (Test_AB.InvokeRequired) {
                Action test_sts_req = delegate { Enable_All_Test_BTN(sts); };
                Test_AB.Invoke(test_sts_req);
            } else {
                Test_AB.Enabled = sts;
                Test_BC.Enabled = sts;
                Test_CB.Enabled = sts;
                Test_BD.Enabled = sts;
                Test_RCD.Enabled = sts;
                Test_diode.Enabled = sts;
                Test_InsulatLine.Enabled = sts;
                Test_InsulatNeut.Enabled = sts;
                Test_ALL_BTN.Enabled = sts;
                Test_PE_open.Enabled = sts;
                Test_diode_open.Enabled = sts;
                cancelBTN.Enabled = !sts && isESP_connected;
                update_btn_Text();
            }
        }

        void update_btn_Text() {
            if (Test_AB.InvokeRequired) {
                Action test_sts_req = delegate { update_btn_Text(); };
                Test_AB.Invoke(test_sts_req);
            } else {
                switch (receive_present_state) {
                    case ReportReceive_states.Standby:
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_InsulatLine.Text = "Test";
                        break;
                    case ReportReceive_states.Req_AB:
                        Test_AB.Text = "Testing...";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_InsulatLine.Text = "Test";
                        break;
                    case ReportReceive_states.Req_BC:
                        Test_BC.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_InsulatLine.Text = "Test";
                        break;
                    case ReportReceive_states.Req_CB:
                        Test_CB.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_InsulatLine.Text = "Test";
                        break;
                    case ReportReceive_states.Req_BD:
                        Test_BD.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_InsulatLine.Text = "Test";
                        break;
                    case ReportReceive_states.Req_Diode:
                        Test_diode.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_RCD.Text = "Test";
                        Test_InsulatLine.Text = "Test";
                        break;
                    case ReportReceive_states.Req_RCD:
                        Test_RCD.Text = "Testing...";
                        Test_AB.Text = "Test";
                        Test_BC.Text = "Test";
                        Test_CB.Text = "Test";
                        Test_BD.Text = "Test";
                        Test_diode.Text = "Test";
                        Test_InsulatLine.Text = "Test";
                        break;
                    case ReportReceive_states.Req_Insul:
                        Test_InsulatLine.Text = "Testing...";
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

        private void addTextToSerialMon(String str) {
            if (SerialMoni.InvokeRequired) {
                Action seri = delegate { addTextToSerialMon(str); };
                SerialMoni.Invoke(seri);
            } else {
                SerialMoni.AppendText(str);
            }

        }


        private void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e) {
            SerialPort sp = (SerialPort)sender;
            //string indata = sp.ReadExisting();
            string indata = "";
            indata = sp.ReadLine();

            addTextToSerialMon(indata);

            switch (receive_present_state) {
                case ReportReceive_states.Standby:
                    break;
                case ReportReceive_states.Req_AB:
                    while (indata.ToCharArray()[0] != '{') {
                        indata = sp.ReadLine();
                    }
                    read_result_A_to_B(JsonSerializer.Deserialize<State_Transition_Test>(indata));
                    testing_Check[0] = true;

                    if (scan_read) {
                        hook_Test_BC();
                        update_btn_Text();
                    } else {
                        receive_present_state = ReportReceive_states.Standby;
                        Enable_All_Test_BTN(true);
                    }

                    break;
                case ReportReceive_states.Req_BC:
                    while (indata.ToCharArray()[0] != '{') {
                        indata = sp.ReadLine();
                    }

                    read_result_B_to_C(JsonSerializer.Deserialize<State_Transition_Test>(indata));
                    testing_Check[1] = true;
                    if (scan_read) {
                        hook_Test_BD();
                        update_btn_Text();
                    } else {
                        receive_present_state = ReportReceive_states.Standby;
                        Enable_All_Test_BTN(true);
                    }

                    break;
                case ReportReceive_states.Req_CB:
                    while (indata.ToCharArray()[0] != '{') {
                        indata = sp.ReadLine();
                    }

                    read_result_C_to_B(JsonSerializer.Deserialize<State_Transition_Test>(indata));
                    testing_Check[2] = true;

                    receive_present_state = ReportReceive_states.Standby;
                    Enable_All_Test_BTN(true);
                    scan_read = false;

                    break;
                case ReportReceive_states.Req_BD://                                                       
                    while (indata.ToCharArray()[0] != '{') {
                        indata = sp.ReadLine();
                    }

                    read_result_B_to_D(JsonSerializer.Deserialize<State_Transition_Test>(indata));
                    testing_Check[3] = true;


                    if (scan_read) {
                        hook_Test_CB();
                        update_btn_Text();
                    } else {
                        receive_present_state = ReportReceive_states.Standby;
                        Enable_All_Test_BTN(true);
                    }

                    break;
                case ReportReceive_states.Req_Diode:
                    while (indata.ToCharArray()[0] != '{') {
                        indata = sp.ReadLine();
                    }

                    read_result_Diode(JsonSerializer.Deserialize<Diode_Test>(indata));
                    testing_Check[4] = true;

                    receive_present_state = ReportReceive_states.Standby;
                    Enable_All_Test_BTN(true);

                    break;
                case ReportReceive_states.Req_RCD:
                    while (indata.ToCharArray()[0] != '{') {
                        indata = sp.ReadLine();
                    }

                    read_result_RCD(JsonSerializer.Deserialize<RCD0>(indata));
                    testing_Check[5] = true;
                    receive_present_state = ReportReceive_states.Standby;
                    Enable_All_Test_BTN(true);
                    break;
                case ReportReceive_states.Req_Insul:
                    while (indata.ToCharArray()[0] != '{') {
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

        void read_result_A_to_B(State_Transition_Test s_A_B) {
            result_State_A_to_B = s_A_B;
            if (AB_PWM_Startup.InvokeRequired) {
                Action add_str = delegate {
                    read_result_A_to_B(s_A_B);
                };
                AB_PWM_Startup.Invoke(add_str);
            } else {
                AB_PWM_Startup.Text = result_State_A_to_B.PWM_StartupDelay + " ms";
            }

            AB_PWM_Amp.Text = result_State_A_to_B.PWM_Amplitude + " V";
            AB_PWM_NVE.Text = result_State_A_to_B.PWM_NveAmplitude + " V";
            AB_PWM_Freq.Text = result_State_A_to_B.PWM_Freq + " Hz";
            AB_PWM_Duty.Text = result_State_A_to_B.PWM_DutyCycle + " %";
            AB_PWM_Imax.Text = result_State_A_to_B.PWM_Imax + " A";



            if (result_State_A_to_B.PWM_StartupDelay_Result) {
                AB_PWM_Startup_Result.Text = "Pass";
                AB_PWM_Startup_Result.ForeColor = Color.Green;
            } else {
                AB_PWM_Startup_Result.Text = "Fail";
                AB_PWM_Startup_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_Amplitude_Result) {
                AB_PWM_Amp_Result.Text = "Pass";
                AB_PWM_Amp_Result.ForeColor = Color.Green;
            } else {
                AB_PWM_Amp_Result.Text = "Fail";
                AB_PWM_Amp_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_NveAmplitude_Result) {
                AB_PWM_NVE_Result.Text = "Pass";
                AB_PWM_NVE_Result.ForeColor = Color.Green;
            } else {
                AB_PWM_NVE_Result.Text = "Fail";
                AB_PWM_NVE_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_Freq_Result) {
                AB_PWM_Freq_Result.Text = "Pass";
                AB_PWM_Freq_Result.ForeColor = Color.Green;
            } else {
                AB_PWM_Freq_Result.Text = "Fail";
                AB_PWM_Freq_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_DutyCycle_Result) {
                AB_PWM_Duty_Result.Text = "Pass";
                AB_PWM_Duty_Result.ForeColor = Color.Green;
            } else {
                AB_PWM_Duty_Result.Text = "Fail";
                AB_PWM_Duty_Result.ForeColor = Color.Red;
            }
            if (result_State_A_to_B.PWM_Imax_Result) {
                AB_PWM_Imax_Result.Text = "Pass";
                AB_PWM_Imax_Result.ForeColor = Color.Green;
            } else {
                AB_PWM_Imax_Result.Text = "Fail";
                AB_PWM_Imax_Result.ForeColor = Color.Red;
            }

            // front page
            B_PWM_Pk.Text = AB_PWM_Amp.Text;
            B_PWM_nPk.Text = AB_PWM_NVE.Text;
            B_PWM_Freq.Text = AB_PWM_Freq.Text;
            B_PWM_Duty.Text = AB_PWM_Duty.Text;
            B_PWM_Imax.Text = AB_PWM_Imax.Text;

            B_PWM_Imax_sts.Text = AB_PWM_Imax_Result.Text;
            B_PWM_Pk_sts.Text = AB_PWM_Amp_Result.Text;
            B_PWM_nPk_sts.Text = AB_PWM_NVE_Result.Text;
            B_PWM_Freq_sts.Text = AB_PWM_Freq_Result.Text;
            B_PWM_Duty_sts.Text = AB_PWM_Duty_Result.Text;

            B_PWM_Imax_sts.ForeColor = AB_PWM_Imax_Result.ForeColor;
            B_PWM_Pk_sts.ForeColor = AB_PWM_Amp_Result.ForeColor;
            B_PWM_nPk_sts.ForeColor = AB_PWM_NVE_Result.ForeColor;
            B_PWM_Freq_sts.ForeColor = AB_PWM_Freq_Result.ForeColor;
            B_PWM_Duty_sts.ForeColor = AB_PWM_Duty_Result.ForeColor;

            update_Freq_AB(result_State_A_to_B.PWM_Freq + " Hz");
            update_Duty_AB(result_State_A_to_B.PWM_DutyCycle + " %");
            update_Vmax_AB(result_State_A_to_B.PWM_Amplitude + " V");
            update_Vmin_AB(result_State_A_to_B.PWM_NveAmplitude + " V");

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
        void read_result_B_to_C(State_Transition_Test s_B_C) {
            result_State_B_to_C = s_B_C;
            if (BC_PWM_Startup.InvokeRequired) {
                Action add_str = delegate {
                    read_result_B_to_C(s_B_C);
                };
                BC_PWM_Startup.Invoke(add_str);
            } else {
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

            if (result_State_B_to_C.PWM_StartupDelay_Result) {
                BC_PWM_Startup_Result.Text = "Pass";
                BC_PWM_Startup_Result.ForeColor = Color.Green;
            } else {
                BC_PWM_Startup_Result.Text = "Fail";
                BC_PWM_Startup_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_Amplitude_Result) {
                BC_PWM_Amp_Result.Text = "Pass";
                BC_PWM_Amp_Result.ForeColor = Color.Green;
            } else {
                BC_PWM_Amp_Result.Text = "Fail";
                BC_PWM_Amp_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_NveAmplitude_Result) {
                BC_PWM_NVE_Result.Text = "Pass";
                BC_PWM_NVE_Result.ForeColor = Color.Green;
            } else {
                BC_PWM_NVE_Result.Text = "Fail";
                BC_PWM_NVE_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_Freq_Result) {
                BC_PWM_Freq_Result.Text = "Pass";
                BC_PWM_Freq_Result.ForeColor = Color.Green;
            } else {
                BC_PWM_Freq_Result.Text = "Fail";
                BC_PWM_Freq_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_DutyCycle_Result) {
                BC_PWM_Duty_Result.Text = "Pass";
                BC_PWM_Duty_Result.ForeColor = Color.Green;
            } else {
                BC_PWM_Duty_Result.Text = "Fail";
                BC_PWM_Duty_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PWM_Imax_Result) {
                BC_PWM_Imax_Result.Text = "Pass";
                BC_PWM_Imax_Result.ForeColor = Color.Green;
            } else {
                BC_PWM_Imax_Result.Text = "Fail";
                BC_PWM_Imax_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.Voltage_Result) {
                BC_Voltage_Result.Text = "Pass";
                BC_Voltage_Result.ForeColor = Color.Green;
            } else {
                BC_Voltage_Result.Text = "Fail";
                BC_Voltage_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.MainsOnDelay_Result) {
                BC_PWM_OnDel_Result.Text = "Pass";
                BC_PWM_OnDel_Result.ForeColor = Color.Green;
            } else {
                BC_PWM_OnDel_Result.Text = "Fail";
                BC_PWM_OnDel_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.MainsFreq_Result) {
                BC_PWM_MainFreq_Result.Text = "Pass";
                BC_PWM_MainFreq_Result.ForeColor = Color.Green;
            } else {
                BC_PWM_MainFreq_Result.Text = "Fail";
                BC_PWM_MainFreq_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_C.PP_Result) {
                BC_PP_Result.Text = "Pass";
                BC_PP_Result.ForeColor = Color.Green;
            } else {
                BC_PP_Result.Text = "Fail";
                BC_PP_Result.ForeColor = Color.Red;
            }

            // Front page
            BC_MainOn.Text = BC_PWM_OnDel.Text;
            BC_MainOn_sts.Text = BC_PWM_OnDel_Result.Text;
            BC_MainOn_sts.ForeColor = BC_PWM_OnDel_Result.ForeColor;

            C_PWM_Pk.Text = BC_PWM_Amp.Text;
            C_PWM_nPk.Text = BC_PWM_NVE.Text;
            C_PWM_Freq.Text = BC_PWM_Freq.Text;
            C_PWM_Duty.Text = BC_PWM_Duty.Text;
            C_PWM_Imax.Text = BC_PWM_Imax.Text;
            C_PWM_MainVolt.Text = BC_Voltage.Text;
            C_PWM_MainFreq.Text = BC_PWM_MainFreq.Text;
            C_PWM_PP.Text = BC_PP.Text;

            C_PWM_Pk_sts.Text = BC_PWM_Amp_Result.Text;
            C_PWM_nPk_sts.Text = BC_PWM_NVE_Result.Text;
            C_PWM_Freq_sts.Text = BC_PWM_Freq_Result.Text;
            C_PWM_Duty_sts.Text = BC_PWM_Duty_Result.Text;
            C_PWM_Imax_sts.Text = BC_PWM_Imax_Result.Text;
            C_PWM_MainVolt_sts.Text = BC_Voltage_Result.Text;
            C_PWM_MainFreq_sts.Text = BC_PWM_MainFreq_Result.Text;
            C_PWM_PP_sts.Text = BC_PP_Result.Text;

            C_PWM_Pk_sts.ForeColor = BC_PWM_Amp_Result.ForeColor;
            C_PWM_nPk_sts.ForeColor = BC_PWM_NVE_Result.ForeColor;
            C_PWM_Freq_sts.ForeColor = BC_PWM_Freq_Result.ForeColor;
            C_PWM_Duty_sts.ForeColor = BC_PWM_Duty_Result.ForeColor;
            C_PWM_Imax_sts.ForeColor = BC_PWM_Imax_Result.ForeColor;
            C_PWM_MainVolt_sts.ForeColor = BC_Voltage_Result.ForeColor;
            C_PWM_MainFreq_sts.ForeColor = BC_PWM_MainFreq_Result.ForeColor;
            C_PWM_PP_sts.ForeColor = BC_PP_Result.ForeColor;

            update_Freq_BC(result_State_B_to_C.PWM_Freq + " Hz");
            update_Duty_BC(result_State_B_to_C.PWM_DutyCycle + " %");
            update_Vmax_BC(result_State_B_to_C.PWM_Amplitude + " V");
            update_Vmin_BC(result_State_B_to_C.PWM_NveAmplitude + " V");
        }
        void read_result_C_to_B(State_Transition_Test s_C_B) {
            result_State_C_to_B = s_C_B;
            if (CB_PWM_Startup.InvokeRequired) {
                Action add_str = delegate {
                    read_result_C_to_B(s_C_B);
                };
                CB_PWM_Startup.Invoke(add_str);
            } else {
                CB_PWM_Startup.Text = result_State_C_to_B.PWM_StartupDelay + " ms";
            }

            CB_PWM_Amp.Text = result_State_C_to_B.PWM_Amplitude + " V";
            CB_PWM_NVE.Text = result_State_C_to_B.PWM_NveAmplitude + " V";
            CB_PWM_Freq.Text = result_State_C_to_B.PWM_Freq + " Hz";
            CB_PWM_Duty.Text = result_State_C_to_B.PWM_DutyCycle + " %";
            CB_PWM_Imax.Text = result_State_C_to_B.PWM_Imax + " A";
            CB_PWM_OffDel.Text = result_State_C_to_B.MainsOffDelay + " ms";

            if (result_State_C_to_B.PWM_StartupDelay_Result) {
                CB_PWM_Startup_Result.Text = "Pass";
                CB_PWM_Startup_Result.ForeColor = Color.Green;
            } else {
                CB_PWM_Startup_Result.Text = "Fail";
                CB_PWM_Startup_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_Amplitude_Result) {
                CB_PWM_Amp_Result.Text = "Pass";
                CB_PWM_Amp_Result.ForeColor = Color.Green;
            } else {
                CB_PWM_Amp_Result.Text = "Fail";
                CB_PWM_Amp_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_NveAmplitude_Result) {
                CB_PWM_NVE_Result.Text = "Pass";
                CB_PWM_NVE_Result.ForeColor = Color.Green;
            } else {
                CB_PWM_NVE_Result.Text = "Fail";
                CB_PWM_NVE_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_Freq_Result) {
                CB_PWM_Freq_Result.Text = "Pass";
                CB_PWM_Freq_Result.ForeColor = Color.Green;
            } else {
                CB_PWM_Freq_Result.Text = "Fail";
                CB_PWM_Freq_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_DutyCycle_Result) {
                CB_PWM_Duty_Result.Text = "Pass";
                CB_PWM_Duty_Result.ForeColor = Color.Green;
            } else {
                CB_PWM_Duty_Result.Text = "Fail";
                CB_PWM_Duty_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.PWM_Imax_Result) {
                CB_PWM_Imax_Result.Text = "Pass";
                CB_PWM_Imax_Result.ForeColor = Color.Green;
            } else {
                CB_PWM_Imax_Result.Text = "Fail";
                CB_PWM_Imax_Result.ForeColor = Color.Red;
            }
            if (result_State_C_to_B.MainsOffDelay_Result) {
                CB_PWM_OffDel_Result.Text = "Pass";
                CB_PWM_OffDel_Result.ForeColor = Color.Green;
            } else {
                CB_PWM_OffDel_Result.Text = "Fail";
                CB_PWM_OffDel_Result.ForeColor = Color.Red;
            }

            CB_PWM_MainOff.Text = CB_PWM_OffDel.Text;
            CB_PWM_MainOff_sts.Text = CB_PWM_OffDel_Result.Text;
            CB_PWM_MainOff_sts.ForeColor = CB_PWM_OffDel_Result.ForeColor;


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
        void read_result_B_to_D(State_Transition_Test s_B_D) {
            result_State_B_to_D = s_B_D;
            if (BD_PWM_Startup.InvokeRequired) {
                Action add_str = delegate {
                    read_result_B_to_D(s_B_D);
                };
                BD_PWM_Startup.Invoke(add_str);
            } else {
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

            if (result_State_B_to_D.PWM_StartupDelay_Result) {
                BD_PWM_Startup_Result.Text = "Pass";
                BD_PWM_Startup_Result.ForeColor = Color.Green;
            } else {
                BD_PWM_Startup_Result.Text = "Fail";
                BD_PWM_Startup_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_Amplitude_Result) {
                BD_PWM_Amp_Result.Text = "Pass";
                BD_PWM_Amp_Result.ForeColor = Color.Green;
            } else {
                BD_PWM_Amp_Result.Text = "Fail";
                BD_PWM_Amp_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_NveAmplitude_Result) {
                BD_PWM_NVE_Result.Text = "Pass";
                BD_PWM_NVE_Result.ForeColor = Color.Green;
            } else {
                BD_PWM_NVE_Result.Text = "Fail";
                BD_PWM_NVE_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_Freq_Result) {
                BD_PWM_Freq_Result.Text = "Pass";
                BD_PWM_Freq_Result.ForeColor = Color.Green;
            } else {
                BD_PWM_Freq_Result.Text = "Fail";
                BD_PWM_Freq_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_DutyCycle_Result) {
                BD_PWM_Duty_Result.Text = "Pass";
                BD_PWM_Duty_Result.ForeColor = Color.Green;
            } else {
                BD_PWM_Duty_Result.Text = "Fail";
                BD_PWM_Duty_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PWM_Imax_Result) {
                BD_PWM_Imax_Result.Text = "Pass";
                BD_PWM_Imax_Result.ForeColor = Color.Green;
            } else {
                BD_PWM_Imax_Result.Text = "Fail";
                BD_PWM_Imax_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.Voltage_Result) {
                BD_Voltage_Result.Text = "Pass";
                BD_Voltage_Result.ForeColor = Color.Green;
            } else {
                BD_Voltage_Result.Text = "Fail";
                BD_Voltage_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.MainsOnDelay_Result) {
                BD_PWM_OnDel_Result.Text = "Pass";
                BD_PWM_OnDel_Result.ForeColor = Color.Green;
            } else {
                BD_PWM_OnDel_Result.Text = "Fail";
                BD_PWM_OnDel_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.MainsFreq_Result) {
                BD_PWM_MainFreq_Result.Text = "Pass";
                BD_PWM_MainFreq_Result.ForeColor = Color.Green;
            } else {
                BD_PWM_MainFreq_Result.Text = "Fail";
                BD_PWM_MainFreq_Result.ForeColor = Color.Red;
            }
            if (result_State_B_to_D.PP_Result) {
                BD_PP_Result.Text = "Pass";
                BD_PP_Result.ForeColor = Color.Green;
            } else {
                BD_PP_Result.Text = "Fail";
                BD_PP_Result.ForeColor = Color.Red;
            }

            // Front page


            D_PWM_Pk.Text = BD_PWM_Amp.Text;
            D_PWM_nPk.Text = BD_PWM_NVE.Text;
            D_PWM_Freq.Text = BD_PWM_Freq.Text;
            D_PWM_Duty.Text = BD_PWM_Duty.Text;
            D_PWM_Imax.Text = BD_PWM_Imax.Text;
            D_PWM_MainVolt.Text = BD_Voltage.Text;
            D_PWM_MainFreq.Text = BD_PWM_MainFreq.Text;
            D_PWM_PP.Text = BD_PP.Text;
            D_PWM_Pk_sts.Text = BD_PWM_Amp_Result.Text;
            D_PWM_nPk_sts.Text = BD_PWM_NVE_Result.Text;
            D_PWM_Freq_sts.Text = BD_PWM_Freq_Result.Text;
            D_PWM_Duty_sts.Text = BD_PWM_Duty_Result.Text;
            D_PWM_Imax_sts.Text = BD_PWM_Imax_Result.Text;
            D_PWM_MainVolt_sts.Text = BD_Voltage_Result.Text;
            D_PWM_MainFreq_sts.Text = BD_PWM_MainFreq_Result.Text;
            D_PWM_PP_sts.Text = BD_PP_Result.Text;
            D_PWM_Pk_sts.ForeColor = BD_PWM_Amp_Result.ForeColor;
            D_PWM_nPk_sts.ForeColor = BD_PWM_NVE_Result.ForeColor;
            D_PWM_Freq_sts.ForeColor = BD_PWM_Freq_Result.ForeColor;
            D_PWM_Duty_sts.ForeColor = BD_PWM_Duty_Result.ForeColor;
            D_PWM_Imax_sts.ForeColor = BD_PWM_Imax_Result.ForeColor;
            D_PWM_MainVolt_sts.ForeColor = BD_Voltage_Result.ForeColor;
            D_PWM_MainFreq_sts.ForeColor = BD_PWM_MainFreq_Result.ForeColor;
            D_PWM_PP_sts.ForeColor = BD_PP_Result.ForeColor;

            update_Freq_CD(result_State_B_to_D.PWM_Freq + " Hz");
            update_Duty_CD(result_State_B_to_D.PWM_DutyCycle + " %");
            update_Vmax_CD(result_State_B_to_D.PWM_Amplitude + " V");
            update_Vmin_CD(result_State_B_to_D.PWM_NveAmplitude + " V");
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

        void read_result_RCD(RCD0 s_rcd) {

            result_rcd_test = s_rcd;

            if (RCD_TripTime.InvokeRequired) {
                Action add_str = delegate {
                    read_result_RCD(s_rcd);
                };
                RCD_TripTime.Invoke(add_str);
            } else {
                RCD_TripTime.Text = result_rcd_test.Trip_Time + " ms";
            }
            RCD_Limit.Text = result_rcd_test.Limit + " ms";
            RCD_Current.Text = result_rcd_test.Current + " mA";

            if (result_rcd_test.RCD0_Result) {
                RCD_check.Text = "Pass";
                RCD_check.ForeColor = Color.Green;
            } else {
                RCD_check.Text = "Fail";
                RCD_check.ForeColor = Color.Red;
            }

            // Front Page

            RCD_TestingInjectedCurrent.Text = RCD_Current.Text;
            RCD_Accecptable_TripTime.Text = RCD_Limit.Text;
            RCD_TripTime_ms.Text = RCD_TripTime.Text;
            RCD_TripTime_ms_sts.Text = RCD_check.Text;
            RCD_TripTime_ms_sts.ForeColor = RCD_check.ForeColor;

        }
        void read_result_Diode(Diode_Test s_diode) {

            result_diode_test = s_diode;

            if (DiodeShort_Delay.InvokeRequired) {
                Action add_str = delegate {
                    read_result_Diode(s_diode);
                };
                DiodeShort_Delay.Invoke(add_str);
            } else {

                if (diode_TestCMD == Diode_Test_ENUM.TestDiode_Short) {
                    diode_TestCMD = Diode_Test_ENUM.notTesting;
                    DiodeShort_Delay.Text = result_diode_test.Diode_ShortCircuit_MainsOffDelay + " ms";
                    if (result_diode_test.Diode_ShortCircuit_Result) {
                        Diode_Short_check.Text = "Pass";
                        Diode_Short_check.ForeColor = Color.Green;
                    } else {
                        Diode_Short_check.Text = "Fail";
                        Diode_Short_check.ForeColor = Color.Red;
                    }
                }


            }

            if (diode_TestCMD == Diode_Test_ENUM.TestPE_Open) {
                diode_TestCMD = Diode_Test_ENUM.notTesting;
                PE_Open_Delay.Text = result_diode_test.PE_OpenCircuit_MainsOffDelay + " ms";
                if (result_diode_test.PE_OpenCircuit_Result) {
                    PE_Open_check.Text = "Pass";
                    PE_Open_check.ForeColor = Color.Green;
                } else {
                    PE_Open_check.Text = "Fail";
                    PE_Open_check.ForeColor = Color.Red;
                }
            }
            if (diode_TestCMD == Diode_Test_ENUM.TestDiode_Open) {
                diode_TestCMD = Diode_Test_ENUM.notTesting;
                DiodeOpen_Delay.Text = result_diode_test.Diode_OpenCircuit_MainsOffDelay + " ms";
                if (result_diode_test.Diode_OpenCircuit_Result) {
                    DiodeOpen_check.Text = "Pass";
                    DiodeOpen_check.ForeColor = Color.Green;
                } else {
                    DiodeOpen_check.Text = "Fail";
                    DiodeOpen_check.ForeColor = Color.Red;
                }
            }

            // Front Page
            DiodeOpen_MainOff.Text = DiodeOpen_Delay.Text;
            PE_Open_MainOff.Text = PE_Open_Delay.Text;
            DiodeSh_MainOff.Text = DiodeShort_Delay.Text;

            DiodeOpen_MainOff_sts.Text = DiodeOpen_check.Text;
            PE_Open_MainOff_sts.Text = PE_Open_check.Text;
            DiodeSh_MainOff_sts.Text = Diode_Short_check.Text;

            DiodeOpen_MainOff_sts.ForeColor = DiodeOpen_check.ForeColor;
            PE_Open_MainOff_sts.ForeColor = PE_Open_check.ForeColor;
            DiodeSh_MainOff_sts.ForeColor = Diode_Short_check.ForeColor;


        }
        void read_result_Insulator(Insulation_Test s_insu) {

            Insulation_Test = s_insu;

            if (Insulator_Limit.InvokeRequired) {
                Action add_str = delegate {
                    read_result_Insulator(s_insu);
                };
                Insulator_Limit.Invoke(add_str);
            } else {
                Insulator_Limit.Text = Insulation_Test.N_PE + " Ω";
            }
            Insulator_Result.Text = Insulation_Test.L_PE + " Ω";
            Insulator_Volt.Text = Insulation_Test.Voltage + " V";

            if (Insulation_Test.Insulation_Testing) {
                Insu_check.Text = "Pass";
                Insu_check.ForeColor = Color.Green;
            } else {
                Insu_check.Text = "Fail";
                Insu_check.ForeColor = Color.Red;
            }

        }

        void read_result_LinePE(Insulation_Test s_insu) {
            Insulation_Test = s_insu;

            if (Insulator_Limit.InvokeRequired) {
                Action add_str = delegate {
                    read_result_Insulator(s_insu);
                };
                Insulator_Limit.Invoke(add_str);
            } else {
                Insulator_Result.Text = Insulation_Test.L_PE + " Ω";
            }

            Insulator_Volt.Text = Insulation_Test.Voltage + " V";

            if (Insulation_Test.Insulation_Testing) {
                Insu_check.Text = "Pass";
                Insu_check.ForeColor = Color.Green;
            } else {
                Insu_check.Text = "Fail";
                Insu_check.ForeColor = Color.Red;
            }
        }
        void read_result_NeutPE(Insulation_Test s_insu) {
            Insulation_Test = s_insu;

            if (Insulator_Limit.InvokeRequired) {
                Action add_str = delegate {
                    read_result_Insulator(s_insu);
                };
                Insulator_Limit.Invoke(add_str);
            } else {
                Insulator_Limit.Text = Insulation_Test.N_PE + " Ω";
            }
            Insulator_Volt.Text = Insulation_Test.Voltage + " V";

            if (Insulation_Test.Insulation_Testing) {
                Insu_check.Text = "Pass";
                Insu_check.ForeColor = Color.Green;
            } else {
                Insu_check.Text = "Fail";
                Insu_check.ForeColor = Color.Red;
            }
        }

        private void Form1_Load(object sender, EventArgs e) {
            clearS_AB();
            clearS_BC();
            clearS_CB();
            clearS_BD();
            clearDiode();
            clearRCD();
            clearInsu();
        }

        private void updateStatus_BTN() {
            //bool en = true;
            bool en = isTemplateSelected;
            /*
            foreach (bool b in testing_Check)
            {
                en &= b;
            }*/


            if (ExportDOCX.InvokeRequired) {
                Action add_str = delegate {
                    updateStatus_BTN();
                };
                ExportDOCX.Invoke(add_str);
            } else {
                // ExportPDF.Enabled = en;
                ExportDOCX.Enabled = en;
            }


        }

        private void Test_AB_Click(object sender, EventArgs e) {
            scan_read = false;

            hook_Test_AB();
            Enable_All_Test_BTN(false);
        }

        private void Test_BC_Click(object sender, EventArgs e) {
            scan_read = false;

            hook_Test_BC();
            Enable_All_Test_BTN(false);
        }

        private void Test_CB_Click(object sender, EventArgs e) {
            scan_read = false;

            hook_Test_CB();
            Enable_All_Test_BTN(false);
        }

        private void Test_BD_Click(object sender, EventArgs e) {
            scan_read = false;
            hook_Test_BD();
            Enable_All_Test_BTN(false);
        }

        private void Test_RCD_Click(object sender, EventArgs e) {
            hook_Test_RCD();
            if (commu_mode == EVSE_Tester_CommunicationMode.SerialPort)
                Enable_All_Test_BTN(false);
        }

        private void Test_Insulat_Click(object sender, EventArgs e) {
            hook_Test_Insulator();
            if (commu_mode == EVSE_Tester_CommunicationMode.SerialPort)
                Enable_All_Test_BTN(false);
        }

        private void button1_Click(object sender, EventArgs e) {
            diode_TestCMD = Diode_Test_ENUM.TestDiode_Short;
            hook_Test_Diode();
            if (commu_mode == EVSE_Tester_CommunicationMode.SerialPort)
                Enable_All_Test_BTN(false);
        }

        private void Test_PE_open_Click(object sender, EventArgs e) {
            diode_TestCMD = Diode_Test_ENUM.TestPE_Open;
            hook_Test_PE_Open();
            if (commu_mode == EVSE_Tester_CommunicationMode.SerialPort)
                Enable_All_Test_BTN(false);
        }

        private void Test_diode_open_Click(object sender, EventArgs e) {
            diode_TestCMD = Diode_Test_ENUM.TestDiode_Open;
            hook_Test_DiodeOpen();
            if (commu_mode == EVSE_Tester_CommunicationMode.SerialPort)
                Enable_All_Test_BTN(false);
        }
        private void TestL_PE_Click(object sender, EventArgs e) {
            hook_Test_LinePE();
            if (commu_mode == EVSE_Tester_CommunicationMode.SerialPort)
                Enable_All_Test_BTN(false);
        }

        private void TestN_PE_Click(object sender, EventArgs e) {
            hook_Test_NeutPE();
            if (commu_mode == EVSE_Tester_CommunicationMode.SerialPort)
                Enable_All_Test_BTN(false);
        }


        private void exp_dir_Click(object sender, EventArgs e) {
            //DialogResult result = Save_to.ShowDialog();

            DialogResult result = SaveReportAs.ShowDialog();
            if (SaveReportAs.FileName.IndexOf(".docx") == -1) // add .docx when no .docx found in name
            {
                SaveReportAs.FileName += ".docx";
            }
            if (result == DialogResult.OK) {
                // Exp_to.Text = Save_to.SelectedPath;
                Exp_to.Text = SaveReportAs.FileName;
            }
        }

        private void Test_ALL_BTN_Click(object sender, EventArgs e) {

            scan_read = true;
            hook_Test_AB();
            // Lock all Test button
            if (commu_mode == EVSE_Tester_CommunicationMode.SerialPort)
                Enable_All_Test_BTN(false);
        }

        private void COM_Input_Enter(object sender, KeyPressEventArgs e) {

        }

        private void COM_Input_KeyDown(object sender, KeyEventArgs e) {
            if (Control.ModifierKeys == Keys.Enter)
                hook_Connect_ESP();
        }

        private void SerialMoni_TextChanged(object sender, EventArgs e) {

        }

        private void cancelBTN_Click(object sender, EventArgs e) {
            receive_present_state = ReportReceive_states.Standby;
            Enable_All_Test_BTN(true);
            scan_read = false;
        }

        void clearS_AB() {
            if (AB_PWM_Startup.InvokeRequired) {
                Action add_str = delegate {
                    clearS_AB();
                };
                AB_PWM_Startup.Invoke(add_str);
            } else {
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
        void clearS_BC() {
            if (BC_PWM_Startup.InvokeRequired) {
                Action add_str = delegate {
                    clearS_BC();
                };
                BC_PWM_Startup.Invoke(add_str);
            } else {
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
        void clearS_CB() {
            if (CB_PWM_Startup.InvokeRequired) {
                Action add_str = delegate {
                    clearS_CB();
                };
                CB_PWM_Startup.Invoke(add_str);
            } else {
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
        void clearS_BD() {
            if (BD_PWM_Startup.InvokeRequired) {
                Action add_str = delegate {
                    clearS_BD();
                };
                BD_PWM_Startup.Invoke(add_str);
            } else {
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
        void clearDiode() {
            if (DiodeShort_Delay.InvokeRequired) {
                Action add_str = delegate {
                    clearDiode();
                };
                DiodeShort_Delay.Invoke(add_str);
            } else {
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
        void clearRCD() {

            if (RCD_TripTime.InvokeRequired) {
                Action add_str = delegate {
                    clearRCD();
                };
                RCD_TripTime.Invoke(add_str);
            } else {
                RCD_TripTime.Text = "-";
            }
            RCD_Limit.Text = "-";
            RCD_Current.Text = "-";

            RCD_check.Text = "(-)";
            RCD_check.ForeColor = Color.Black;

        }
        void clearInsu() {
            if (Insulator_Limit.InvokeRequired) {
                Action add_str = delegate {
                    clearInsu();
                };
                Insulator_Limit.Invoke(add_str);
            } else {
                Insulator_Limit.Text = "-";
            }
            Insulator_Result.Text = "-";
            Insulator_Volt.Text = "-";

            Insu_check.Text = "(-)";
            Insu_check.ForeColor = Color.Black;

        }

        private void ClrResult_Click(object sender, EventArgs e) {
            clearS_AB();
            clearS_BC();
            clearS_CB();
            clearS_BD();
            clearDiode();
            clearRCD();
            clearInsu();
        }

        private void ExportDOCX_Click(object sender, EventArgs e) {
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
            if (result == DialogResult.OK) {
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

                if (result_diode_test.Diode_ShortCircuit_Result) { diode_sh_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { diode_sh_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_diode_test.PE_OpenCircuit_Result) { pe_op_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { pe_op_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }
                if (result_diode_test.Diode_OpenCircuit_Result) { diode_op_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { diode_op_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                MiniWordColorText rcd_testing_result;
                if (result_rcd_test.RCD0_Result) { rcd_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { rcd_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }

                MiniWordColorText insulation_testing_result;
                if (Insulation_Test.Insulation_Testing) { insulation_testing_result = new MiniWordColorText { Text = "Pass", FontColor = "#50C878" }; } else { insulation_testing_result = new MiniWordColorText { Text = "Fail", FontColor = "#660000" }; }



                var value = new Dictionary<string, object>() {
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

        private void SelTelp_Click(object sender, EventArgs e) {
            DialogResult result = WordTemplate.ShowDialog();
            if (result == DialogResult.OK) {
                ExportDOCX.Text = "Export Report DOCX";
                ExportDOCX.Enabled = true;
                //Exp_to.Text = Save_to.SelectedPath;
                Templ_from.Text = WordTemplate.FileName;
            }
        }

        private void Save_to_HelpRequest(object sender, EventArgs e) {

        }

        private void Select_File_word(object sender, CancelEventArgs e) {
            isTemplateSelected = true;
            ExportDOCX.Text = "Export Report DOCX";
            ExportDOCX.Enabled = true;
            //Exp_to.Text = Save_to.SelectedPath;
            Templ_from.Text = WordTemplate.FileName;
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e) {

        }

        private void onDropdown(object sender, EventArgs e) {
            scanBluetooth();
        }

        private void scanBluetooth() {
            Bluetooth_Devices_List.Items.Clear();
            // Scan Bluetooth
            devices = client.DiscoverDevicesInRange();

            foreach (BluetoothDeviceInfo deviceI in devices) {
                Bluetooth_Devices_List.Items.Add(deviceI.DeviceName);
            }
        }

        private void onsel_Bluetooth(object sender, EventArgs e) {
            //indx_sel.Text = Bluetooth_Devices_List.SelectedIndex.ToString();
            try {
                Selected_Device = devices[Bluetooth_Devices_List.SelectedIndex];
                Bluetooth_Connect.Enabled = true;
                Bluetooth_Devices_List.Text = Selected_Device.DeviceName;
            } catch (IndexOutOfRangeException ex) {

            }

        }

        private void Bluetooth_Connect_Click(object sender, EventArgs e) {
            if (!bluetooth_connecting) {
                try {
                    Bluetooth_Devices_List.Enabled = false;
                    Bluetooth_Connect.Text = "Connecting...";
                } catch (Exception) {

                } finally {
                    try {
                        BluetoothEndPoint endPoint = new BluetoothEndPoint(Selected_Device.DeviceAddress, BluetoothService.SerialPort);
                        client.Connect(endPoint);
                        bluetoothStream = client.GetStream();
                        bluetoothStream.ReadTimeout = 30000; // set read timeout to 5 sec

                        bluetooth_connecting = true;
                        Bluetooth_Connect.Text = "Disconnect";

                        COM_Input.Enabled = false;
                        Connect_ESP_BTN.Enabled = false;

                        commu_mode = EVSE_Tester_CommunicationMode.Bluetooth;
                        updateESP32_Connection_Status(true);
                        isESP_connected = true;

                        TestingMode.Visible = true;

                    } catch (Exception) {
                        Bluetooth_Connect.Text = "Connect";
                        Bluetooth_Devices_List.Enabled = true;
                        updateESP32_Connection_Status(false);
                        isESP_connected = false;
                        COM_Input.Enabled = true;
                        Connect_ESP_BTN.Enabled = true;

                        TestingMode.Visible = false;

                        commu_mode = EVSE_Tester_CommunicationMode.None;
                    }
                }
            } else {
                Bluetooth_Connect.Text = "Connect";
                Bluetooth_Devices_List.Enabled = true;
                client.Close();
                bluetoothStream.Close();


                client = new BluetoothClient();
                bluetoothStream = null;

                bluetooth_connecting = false;
                updateESP32_Connection_Status(false);
                isESP_connected = false;
                COM_Input.Enabled = true;
                Connect_ESP_BTN.Enabled = true;
                commu_mode = EVSE_Tester_CommunicationMode.None;
            }



        }

        private void TestBluetooth_Click(object sender, EventArgs e) {
            if (bluetoothStream.CanWrite) {
                string message = "Hello, Bluetooth!";
                byte[] messageBuffer = Encoding.ASCII.GetBytes(message);
                bluetoothStream.Write(messageBuffer, 0, messageBuffer.Length);

                byte[] receiveBuffer = new byte[512];
                bluetoothStream.Read(receiveBuffer, 0, receiveBuffer.Length);
                TestBluetoothTxt.Text = Encoding.ASCII.GetString(receiveBuffer);
            }
        }

        private void send_Bluetooth(string str) {
            try {
                byte[] msg = Encoding.ASCII.GetBytes(str);
                bluetoothStream.Write(msg, 0, msg.Length);
                bluetoothStream.Flush();
            } catch (IOException ex) {
                Bluetooth_Connect.Text = "Connect";
                Bluetooth_Devices_List.Enabled = true;
                updateESP32_Connection_Status(false);
                isESP_connected = false;
                COM_Input.Enabled = true;
                Connect_ESP_BTN.Enabled = true;

                commu_mode = EVSE_Tester_CommunicationMode.None;
            }

        }
        private string receive_Bluetooth() {
            string indata = "";
            int rem_;
            char lastCh;
            bool endJson = false;
            do {
                try {
                    rem_ = bluetoothStream.Read(receiveBuffer, 0, receiveBuffer.Length);
                    if (rem_ != 0) {
                        lastCh = (char)receiveBuffer[rem_ - 1];
                        if (lastCh == '}') {
                            endJson = true;
                        } else {
                            for (UInt32 i = 0; i < rem_; i++) {
                                if ((char)receiveBuffer[i] == '}') {
                                    endJson = true;
                                    break;
                                }
                            }
                        }
                        indata += Encoding.ASCII.GetString(receiveBuffer, 0, rem_);
                    } else {
                        endJson = true;
                    }
                } catch (IOException e) {

                }


            } while (!endJson);


            addTextToSerialMon(indata);
            return indata;
        }
        private string receive_Array_Bluetooth() {
            string indata = "";
            int rem_;
            char lastCh;
            bool endJson = false;
            do {
                try {
                    rem_ = bluetoothStream.Read(receiveBuffer, 0, receiveBuffer.Length);
                    if (rem_ != 0) {
                        lastCh = (char)receiveBuffer[rem_ - 1];
                        if (lastCh == ']') {
                            endJson = true;
                        } else {
                            for (UInt32 i = 0; i < rem_; i++) {
                                if ((char)receiveBuffer[i] == ']') {
                                    endJson = true;
                                    break;
                                }
                            }
                        }
                        indata += Encoding.ASCII.GetString(receiveBuffer, 0, rem_);
                    } else {
                        endJson = true;
                    }
                } catch (IOException e) {

                }


            } while (!endJson);

            return indata;
        }


        private void INJ_readCP_Click(object sender, EventArgs e) {
            send_Bluetooth("read_CP\n");
        }

        private void INJ_readPP_Click(object sender, EventArgs e) {
            send_Bluetooth("read_PP\n");
        }

        private void INJ_readINS_Click(object sender, EventArgs e) {
            send_Bluetooth("read_INS\n");
        }

        private void ManualTestBTN_Click(object sender, EventArgs e) {
            switch (manual_State) {
                case EVSE_Manual_State.EV_State_A:
                    break;
                case EVSE_Manual_State.EV_State_B:
                    scan_read = false;
                    hook_Test_AB();
                    break;
                case EVSE_Manual_State.EV_State_C:
                    scan_read = false;
                    hook_Test_BC();
                    break;
                case EVSE_Manual_State.EV_State_D:
                    scan_read = false;
                    hook_Test_BD();
                    break;
            }
        }

        void updateManualState() {
            switch (manual_State) {
                case EVSE_Manual_State.EV_State_A:
                    selectState_A.Enabled = true;
                    selectState_B.Enabled = true;
                    selectState_C.Enabled = false;
                    selectState_D.Enabled = false;
                    break;
                case EVSE_Manual_State.EV_State_B:
                    selectState_A.Enabled = true;
                    selectState_B.Enabled = true;
                    selectState_C.Enabled = true;
                    selectState_D.Enabled = false;
                    break;
                case EVSE_Manual_State.EV_State_C:
                    selectState_A.Enabled = false;
                    selectState_B.Enabled = true;
                    selectState_C.Enabled = true;
                    selectState_D.Enabled = true;
                    break;
                case EVSE_Manual_State.EV_State_D:
                    selectState_A.Enabled = false;
                    selectState_B.Enabled = false;
                    selectState_C.Enabled = true;
                    selectState_D.Enabled = true;
                    break;
            }
        }


        private void TestingMode_SelectedIndexChanged(object sender, EventArgs e) {
            if (TestingMode.Text == "Manual") {
                ManualTestBTN.Visible = true;
                selectState_A.Visible = true;
                selectState_B.Visible = true;
                selectState_C.Visible = true;
                selectState_D.Visible = true;

                Test_ALL_BTN.Visible = false;
                cancelBTN.Visible = false;

                AutoTest_Group.Visible = false;
                ManualTest_Group.Visible = true;
                send_Bluetooth("gotoManual");
                manual_State = EVSE_Manual_State.EV_State_A;
                updateManualState();
                selectState_A.Checked = true;
            } else if (TestingMode.Text == "Auto") {
                ManualTestBTN.Visible = false;
                selectState_A.Visible = false;
                selectState_B.Visible = false;
                selectState_C.Visible = false;
                selectState_D.Visible = false;

                Test_ALL_BTN.Visible = true;
                cancelBTN.Visible = true;

                AutoTest_Group.Visible = true;
                ManualTest_Group.Visible = false;
                send_Bluetooth("gotoAuto");
            }
        }

        private void selectState_A_Click(object sender, EventArgs e) {
            manual_State = EVSE_Manual_State.EV_State_A;
            updateManualState();
            send_Bluetooth("Force_A");
        }

        private void selectState_B_Click(object sender, EventArgs e) {
            manual_State = EVSE_Manual_State.EV_State_B;
            updateManualState();
            send_Bluetooth("Force_B");
        }

        private void selectState_C_Click(object sender, EventArgs e) {
            manual_State = EVSE_Manual_State.EV_State_C;
            updateManualState();
            send_Bluetooth("Force_C");
        }

        private void selectState_D_Click(object sender, EventArgs e) {
            manual_State = EVSE_Manual_State.EV_State_D;
            updateManualState();
            send_Bluetooth("Force_D");
        }

        private void AutoScheme_SelectedIndexChanged(object sender, EventArgs e) {
            switch (AutoScheme.SelectedIndex) {
                case 0:
                    picScheme.Image = Properties.Resources.TestingScheme1;
                    break;
                case 1:
                    picScheme.Image = Properties.Resources.TestingSCheme2;
                    break;
            }
        }

        private void picScheme_Click(object sender, EventArgs e) {

        }

        private void LoopTimer_Tick(object sender, EventArgs e) {

        }

        private void testingGraphic_Click(object sender, EventArgs e) {

        }

        private void SamplingCP_Click(object sender, EventArgs e) {
            //CP_Sampling
            send_Bluetooth("CP_Sampling");

            //Test_CP_Sample
            //Test_CP_Sample.Text = receive_Array_Bluetooth();
            String jsonArray = receive_Array_Bluetooth();
            float[] cp_array_sample = JsonSerializer.Deserialize<float[]>(jsonArray);
            Test_CP_Sample.Text = cp_array_sample.Length.ToString();
            //Waveform_pic
            Waveform_pic.Image = DrawOscilloscope(cp_array_sample, Waveform_pic.Width, Waveform_pic.Height, 50000.0f);
        }

        public static Bitmap DrawOscilloscope(float[] samples, int width, int height, float samplingRate) {
            Bitmap bmp = new Bitmap(width, height);
            using Graphics g = Graphics.FromImage(bmp);

            // Styling
            g.Clear(Color.Black);
            Pen gridPen = new Pen(Color.FromArgb(40, 255, 255, 255), 1);
            Pen axisPen = new Pen(Color.Gray, 1);
            Pen waveformPen = new Pen(Color.Lime, 1.5f);
            Font labelFont = new Font("Arial", 10);
            Brush labelBrush = Brushes.White;

            int margin = 50;
            int plotWidth = width - margin * 2;
            int plotHeight = height - margin * 2;

            // Find min/max of samples
            float min = float.MaxValue, max = float.MinValue;
            //foreach (var s in samples)
            //{
            //    if (s < min) min = s;
            //    if (s > max) max = s;
            //}

            max = 15.0f;
            min = -15.0f;

            float range = max - min;
            if (range == 0) range = 1;

            // Total duration in microseconds
            float totalDurationUs = samples.Length * 1000000.0f / samplingRate;

            // Horizontal grid lines & Y-axis labels
            int gridY = 10;
            for (int i = 0; i <= gridY; i++) {
                float y = margin + i * plotHeight / gridY;
                g.DrawLine(gridPen, margin, y, width - margin, y);
                float value = max - i * range / gridY;
                g.DrawString($"{value:0.000}", labelFont, labelBrush, 5, y - 8);
            }

            // Vertical grid lines & X-axis labels
            int gridX = 20;  // 11 grid lines → 10 divisions (adjust if needed)
            for (int i = 0; i <= gridX; i++) {
                float x = margin + i * plotWidth / gridX;
                g.DrawLine(gridPen, x, margin, x, height - margin);
                if ((i % 2) == 0) {
                    float timeUs = i * totalDurationUs / gridX;
                    g.DrawString($"{timeUs:0} µs", labelFont, labelBrush, x - 20, height - margin + 5);
                }

            }

            // Draw bounding box
            g.DrawRectangle(axisPen, margin, margin, plotWidth, plotHeight);

            // Draw waveform scaled to fit
            float xScale = (float)plotWidth / (samples.Length - 1);
            float yScale = (float)plotHeight / range;

            for (int i = 0; i < samples.Length - 1; i++) {
                float x1 = margin + i * xScale;
                float y1 = margin + (max - samples[i]) * yScale;
                float x2 = margin + (i + 1) * xScale;
                float y2 = margin + (max - samples[i + 1]) * yScale;

                g.DrawLine(waveformPen, x1, y1, x2, y2);
            }

            return bmp;
        }

        public static Bitmap DrawFFT_Signal(float[] samples, int width, int height, float samplingRate, double freqMin, double freqMax) {
            Bitmap bmp = new Bitmap(width, height);
            using Graphics g = Graphics.FromImage(bmp);
            // Styling
            g.Clear(Color.Black);
            Pen gridPen = new Pen(Color.FromArgb(40, 255, 255, 255), 1);
            Pen axisPen = new Pen(Color.Gray, 1);
            Pen fftPen = new Pen(Color.Cyan, 1.5f);
            Font labelFont = new Font("Arial", 10);
            Brush labelBrush = Brushes.White;

            int margin = 50;
            int plotWidth = width - margin * 2;
            int plotHeight = height - margin * 2;

            // Zero-padding to next power of 2
            int fftSize = 1;
            while (fftSize < samples.Length) fftSize <<= 1;

            Complex[] fftBuffer = new Complex[fftSize];
            for (int i = 0; i < samples.Length; i++)
                fftBuffer[i] = new Complex(samples[i], 0);

            // Perform FFT (or use DFT_Precise)
            FFT(fftBuffer);

            // Get magnitude spectrum (first half only)
            int spectrumSize = fftSize / 2;
            double[] magnitudes = new double[spectrumSize];
            for (int i = 0; i < spectrumSize; i++) {
                magnitudes[i] = fftBuffer[i].Magnitude;
                magnitudes[i] /= fftSize / 2;
            }

            // Frequency bin resolution
            double binResolution = samplingRate / fftSize;

            // Determine bin range for desired frequency window
            int startBin = (int)(freqMin / binResolution);
            int endBin = (int)(freqMax / binResolution);
            if (endBin >= spectrumSize) endBin = spectrumSize - 1;

            int displayBins = endBin - startBin + 1;
            if (displayBins <= 1) displayBins = 2;

            // Get max magnitude in the selected range
            double maxMag = magnitudes.Skip(startBin).Take(displayBins).Max();
            double minMag = 0;
            double rangeMag = maxMag - minMag;
            if (rangeMag == 0) rangeMag = 1;

            // Horizontal grid lines & Y-axis labels (magnitude)
            int gridY = 10;
            for (int i = 0; i <= gridY; i++) {
                float y = margin + i * plotHeight / gridY;
                g.DrawLine(gridPen, margin, y, width - margin, y);
                double value = maxMag - i * rangeMag / gridY;
                g.DrawString($"{value:0.00}", labelFont, labelBrush, 5, y - 8);
            }

            // Vertical grid lines & X-axis labels (frequency)
            int gridX = 10;
            for (int i = 0; i <= gridX; i++) {
                float x = margin + i * plotWidth / gridX;
                g.DrawLine(gridPen, x, margin, x, height - margin);
                double freq = freqMin + i * (freqMax - freqMin) / gridX;
                g.DrawString($"{freq:0} Hz", labelFont, labelBrush, x - 20, height - margin + 5);
            }

            // Draw bounding box
            g.DrawRectangle(axisPen, margin, margin, plotWidth, plotHeight);

            // Draw FFT spectrum
            float xScale = (float)plotWidth / (displayBins - 1);
            float yScale = (float)plotHeight / (float)rangeMag;

            for (int i = startBin; i < endBin; i++) {
                float x1 = margin + (i - startBin) * xScale;
                float y1 = margin + (float)((maxMag - magnitudes[i]) * yScale);
                float x2 = margin + (i + 1 - startBin) * xScale;
                float y2 = margin + (float)((maxMag - magnitudes[i + 1]) * yScale);

                g.DrawLine(fftPen, x1, y1, x2, y2);
            }

            return bmp;

        }

        private static void FFT(Complex[] buffer) {
            int n = buffer.Length;
            int bits = (int)Math.Log2(n);

            // Bit reversal
            for (int i = 0; i < n; i++) {
                int j = BitReverse(i, bits);
                if (j > i) {
                    var temp = buffer[i];
                    buffer[i] = buffer[j];
                    buffer[j] = temp;
                }
            }

            for (int len = 2; len <= n; len <<= 1) {
                double angle = -2 * Math.PI / len;
                Complex wLen = new Complex(Math.Cos(angle), Math.Sin(angle));
                for (int i = 0; i < n; i += len) {
                    Complex w = Complex.One;
                    for (int j = 0; j < len / 2; j++) {
                        Complex u = buffer[i + j];
                        Complex v = buffer[i + j + len / 2] * w;
                        buffer[i + j] = u + v;
                        buffer[i + j + len / 2] = u - v;
                        w *= wLen;
                    }
                }
            }
        }
        private static int BitReverse(int n, int bits) {
            int reversed = 0;
            for (int i = 0; i < bits; i++) {
                reversed <<= 1;
                reversed |= (n & 1);
                n >>= 1;
            }
            return reversed;
        }

        private void FFT_BTN_Click(object sender, EventArgs e) {
            // Square wave generation
            float[] signal = new float[16000];
            float samplingRate = 4000000.0f;
            Random rnd = new Random();
            float mag = rnd.NextSingle() * 12.0f;

            int samplesPerPeriod = (int)(samplingRate / 1000.0);       // 4000 samples
            int halfPeriod = samplesPerPeriod / 2;                     // 2000 samples

            for (int i = 0; i < signal.Length; i++) {
                signal[i] = (i % samplesPerPeriod < halfPeriod) ? mag : -mag;
            }
            //for(int i = 0;i < 4; i++) {
            //    for (int j = 0; j < 50; j++) {
            //        if(j < 25)
            //        {
            //            signal[50 * i + j] = mag;
            //        }
            //        else
            //        {
            //            signal[50 * i + j] = -mag;
            //        }
            //    }
            //}

            //CP_B2_Pic.Image = DrawFFT_Signal(signal,CP_B2_Pic.Width,CP_B2_Pic.Height,samplingRate);
            using (Bitmap fftImage = DrawFFT_Signal(signal, CP_B2_Pic.Width, CP_B2_Pic.Height, samplingRate, 0, 10000)) {
                CP_B2_Pic.Image?.Dispose(); // Dispose old image if replacing
                CP_B2_Pic.Image = new Bitmap(fftImage); // Clone if needed
            }
        }

        private void TmeDomain_Click(object sender, EventArgs e) {
            // Square wave generation
            float[] signal = new float[16000];
            float samplingRate = 4000000.0f;
            Random rnd = new Random();
            float mag = rnd.NextSingle() * 12.0f;

            int samplesPerPeriod = (int)(samplingRate / 1000.0);       // 4000 samples
            int halfPeriod = samplesPerPeriod / 2;                     // 2000 samples

            for (int i = 0; i < signal.Length; i++) {
                signal[i] = (i % samplesPerPeriod < halfPeriod) ? mag : -mag;
            }
            //for(int i = 0;i < 4; i++) {
            //    for (int j = 0; j < 50; j++) {
            //        if(j < 25)
            //        {
            //            signal[50 * i + j] = mag;
            //        }
            //        else
            //        {
            //            signal[50 * i + j] = -mag;
            //        }
            //    }
            //}

            //CP_B2_Pic.Image = DrawFFT_Signal(signal,CP_B2_Pic.Width,CP_B2_Pic.Height,samplingRate);
            using (Bitmap fftImage = DrawOscilloscope(signal, CP_B2_Pic.Width, CP_B2_Pic.Height, samplingRate)) {
                CP_B2_Pic.Image?.Dispose(); // Dispose old image if replacing
                CP_B2_Pic.Image = new Bitmap(fftImage); // Clone if needed
            }
        }

        private void AB_Freq_CheckedChanged(object sender, EventArgs e) {
            if (AB_Freq.Checked)
                CP_B2_Pic.Image = DrawFFT_Signal(cp_sample_AB, CP_B2_Pic.Width, CP_B2_Pic.Height, 50000.0f, 0.0, 25000.0);
        }

        private void AB_Time_CheckedChanged(object sender, EventArgs e) {
            if (AB_Time.Checked)
                CP_B2_Pic.Image = DrawOscilloscope(cp_sample_AB, CP_B2_Pic.Width, CP_B2_Pic.Height, 50000.0f);

        }

        private void BC_Time_CheckedChanged(object sender, EventArgs e) {
            if (BC_Time.Checked)
                CP_C2_Pic.Image = DrawOscilloscope(cp_sample_BC, CP_C2_Pic.Width, CP_C2_Pic.Height, 50000.0f);
        }

        private void BC_Freq_CheckedChanged(object sender, EventArgs e) {
            if (BC_Freq.Checked)
                CP_C2_Pic.Image = DrawFFT_Signal(cp_sample_BC, CP_C2_Pic.Width, CP_C2_Pic.Height, 50000.0f, 0.0, 25000.0);
        }

        private void CD_Time_CheckedChanged(object sender, EventArgs e) {
            if (CD_Time.Checked)
                CP_D_Pic.Image = DrawOscilloscope(cp_sample_CD, CP_D_Pic.Width, CP_D_Pic.Height, 50000.0f);
        }

        private void CD_Freq_CheckedChanged(object sender, EventArgs e) {
            if (CD_Freq.Checked)
                CP_D_Pic.Image = DrawFFT_Signal(cp_sample_CD, CP_D_Pic.Width, CP_D_Pic.Height, 50000.0f, 0.0, 25000.0);
        }

        private void update_Freq_AB(string fAB) {
            if (Freq_AB.InvokeRequired) {
                Action action = delegate {
                    Freq_AB.Text = fAB;
                };
                Freq_AB.Invoke(action);
            } else {
                Freq_AB.Text = fAB;
            }
        }
        private void update_Duty_AB(string dAB) {
            if (Duty_AB.InvokeRequired) {
                Action action = delegate {
                    Duty_AB.Text = dAB;
                };
                Duty_AB.Invoke(action);
            } else {
                Duty_AB.Text = dAB;
            }
        }
        private void update_Vmax_AB(string vmax) {
            if (Vmax_AB.InvokeRequired) {
                Action action = delegate {
                    Vmax_AB.Text = vmax;
                };
                Vmax_AB.Invoke(action);
            } else {
                Vmax_AB.Text = vmax;
            }
        }
        private void update_Vmin_AB(string vmin) {
            if (Vmin_AB.InvokeRequired) {
                Action action = delegate {
                    Vmin_AB.Text = vmin;
                };
                Vmin_AB.Invoke(action);
            } else {
                Vmin_AB.Text = vmin;
            }
        }
        private void update_Freq_BC(string fAB) {
            if (Freq_BC.InvokeRequired) {
                Action action = delegate {
                    Freq_BC.Text = fAB;
                };
                Freq_BC.Invoke(action);
            } else {
                Freq_BC.Text = fAB;
            }
        }
        private void update_Duty_BC(string dAB) {
            if (Duty_BC.InvokeRequired) {
                Action action = delegate {
                    Duty_BC.Text = dAB;
                };
                Duty_BC.Invoke(action);
            } else {
                Duty_BC.Text = dAB;
            }
        }
        private void update_Vmax_BC(string vmax) {
            if (Vmax_BC.InvokeRequired) {
                Action action = delegate {
                    Vmax_BC.Text = vmax;
                };
                Vmax_BC.Invoke(action);
            } else {
                Vmax_BC.Text = vmax;
            }
        }
        private void update_Vmin_BC(string vmin) {
            if (Vmin_BC.InvokeRequired) {
                Action action = delegate {
                    Vmin_BC.Text = vmin;
                };
                Vmin_BC.Invoke(action);
            } else {
                Vmin_BC.Text = vmin;
            }
        }
        private void update_Freq_CD(string fAB) {
            if (Freq_CD.InvokeRequired) {
                Action action = delegate {
                    Freq_CD.Text = fAB;
                };
                Freq_CD.Invoke(action);
            } else {
                Freq_CD.Text = fAB;
            }
        }
        private void update_Duty_CD(string dAB) {
            if (Duty_CD.InvokeRequired) {
                Action action = delegate {
                    Duty_CD.Text = dAB;
                };
                Duty_CD.Invoke(action);
            } else {
                Duty_CD.Text = dAB;
            }
        }
        private void update_Vmax_CD(string vmax) {
            if (Vmax_CD.InvokeRequired) {
                Action action = delegate {
                    Vmax_CD.Text = vmax;
                };
                Vmax_CD.Invoke(action);
            } else {
                Vmax_CD.Text = vmax;
            }
        }
        private void update_Vmin_CD(string vmin) {
            if (Vmin_CD.InvokeRequired) {
                Action action = delegate {
                    Vmin_CD.Text = vmin;
                };
                Vmin_CD.Invoke(action);
            } else {
                Vmin_CD.Text = vmin;
            }
        }

        private void button1_Click_1(object sender, EventArgs e) {
            send_Bluetooth("AVR_reset\n"); 
        }
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

    enum EVSE_Tester_CommunicationMode
    {
        None = 0,
        SerialPort = 1,
        Bluetooth = 2
    }

    enum Diode_Test_ENUM { 
        notTesting,
        TestDiode_Short,
        TestPE_Open,
        TestDiode_Open
    }
    enum EVSE_Manual_State { 
        EV_State_A,
        EV_State_B,
        EV_State_C,
        EV_State_D
    }
}