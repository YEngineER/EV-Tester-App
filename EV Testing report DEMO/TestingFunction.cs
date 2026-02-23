using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;

namespace EV_Testing_report_DEMO {
    public partial class Form1 : Form {

        private void hook_Test_AB() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_A_to_B\n");
                    receive_present_state = ReportReceive_states.Req_AB;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    if (scan_read) {
                        send_Bluetooth("State_A_to_B\n");
                    } else {
                        send_Bluetooth("State_A_to_B_Single\n");
                    }
                    State_Transition_Test state_Transition_Test = JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth());
                    //read_result_A_to_B(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));
                    read_result_A_to_B(state_Transition_Test);
                    //CP_Sampling
                    send_Bluetooth("CP_Sampling");
                    String jsonArray = receive_Array_Bluetooth();
                    //float[] cp_array_sample = JsonSerializer.Deserialize<float[]>(jsonArray);
                    cp_sample_AB = JsonSerializer.Deserialize<float[]>(jsonArray);
                    //Test_CP_Sample.Text = cp_array_sample.Length.ToString();
                    //Waveform_pic
                    //CP_B2_Pic.Image = DrawOscilloscope(cp_sample_AB, CP_B2_Pic.Width, CP_B2_Pic.Height, 50000.0f);
                    AB_Time.Checked = true;

                    if (scan_read) {
                        hook_Test_BC();
                    }

                    break;
            }
        }
        private void hook_Test_BC() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_B_to_C\n");
                    receive_present_state = ReportReceive_states.Req_BC;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    if (scan_read) {
                        send_Bluetooth("State_B_to_C\n");
                    } else {
                        send_Bluetooth("State_B_to_C_Single\n");
                    }

                    read_result_B_to_C(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));
                    //CP_Sampling
                    send_Bluetooth("CP_Sampling");
                    String jsonArray = receive_Array_Bluetooth();
                    //float[] cp_array_sample = JsonSerializer.Deserialize<float[]>(jsonArray);
                    cp_sample_BC = JsonSerializer.Deserialize<float[]>(jsonArray);
                    //Test_CP_Sample.Text = cp_array_sample.Length.ToString();
                    //Waveform_pic
                    BC_Time.Checked = true;
                    //CP_C2_Pic.Image = DrawOscilloscope(cp_sample_BC, CP_C2_Pic.Width, CP_C2_Pic.Height, 50000.0f);
                    if (scan_read) {
                        hook_Test_BD();
                    }

                    break;
            }

        }
        private void hook_Test_CB() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_C_to_B\n");
                    receive_present_state = ReportReceive_states.Req_CB;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    if (scan_read) {
                        send_Bluetooth("State_C_to_B\n");
                    } else {
                        send_Bluetooth("State_C_to_B_Single\n");
                    }

                    read_result_C_to_B(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));

                    Enable_All_Test_BTN(true);
                    scan_read = false;
                    break;
            }

        }
        private void hook_Test_BD() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_B_to_D\n");
                    receive_present_state = ReportReceive_states.Req_BD;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    if (scan_read) {
                        send_Bluetooth("State_B_to_D\n");
                    } else {
                        send_Bluetooth("State_B_to_D_Single\n");
                    }

                    read_result_B_to_D(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));
                    //CP_Sampling
                    send_Bluetooth("CP_Sampling");
                    String jsonArray = receive_Array_Bluetooth();
                    //float[] cp_array_sample = JsonSerializer.Deserialize<float[]>(jsonArray);
                    cp_sample_CD = JsonSerializer.Deserialize<float[]>(jsonArray);
                    //Test_CP_Sample.Text = cp_array_sample.Length.ToString();
                    //Waveform_pic
                    //CP_D_Pic.Image = DrawOscilloscope(cp_sample_CD, CP_D_Pic.Width, CP_D_Pic.Height, 50000.0f);
                    CD_Time.Checked = true;
                    if (scan_read) {
                        hook_Test_CB();
                    }

                    break;
            }

        }
        private void hook_Test_RCD() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("RCD0_Test\n");
                    receive_present_state = ReportReceive_states.Req_RCD;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("RCD0_Test\n");

                    read_result_RCD(JsonSerializer.Deserialize<RCD0>(receive_Bluetooth()));
                    break;
            }

        }
        private void hook_Test_Insulator() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Insulator_Test\n");
                    receive_present_state = ReportReceive_states.Req_Insul;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("Insulator_Test\n");

                    read_result_Insulator(JsonSerializer.Deserialize<Insulation_Test>(receive_Bluetooth()));
                    break;
            }

        }
        private void hook_Test_LinePE() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Test_LinePE\n");
                    receive_present_state = ReportReceive_states.Req_Insul;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("Test_LinePE\n");

                    read_result_LinePE(JsonSerializer.Deserialize<Insulation_Test>(receive_Bluetooth()));
                    break;
            }
        }
        private void hook_Test_NeutPE() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Test_NeutralPE\n");
                    receive_present_state = ReportReceive_states.Req_Insul;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("Test_NeutralPE\n");

                    read_result_NeutPE(JsonSerializer.Deserialize<Insulation_Test>(receive_Bluetooth()));
                    break;
            }
        }
        private void hook_Test_Diode() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Diode_Test\n");
                    receive_present_state = ReportReceive_states.Req_Diode;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("Diode_Test\n");

                    read_result_Diode(JsonSerializer.Deserialize<Diode_Test>(receive_Bluetooth()));
                    break;
            }

        }
        private void hook_Test_PE_Open() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("PE_Open_Test\n");
                    receive_present_state = ReportReceive_states.Req_Diode;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("PE_Open_Test\n");

                    read_result_Diode(JsonSerializer.Deserialize<Diode_Test>(receive_Bluetooth()));
                    break;
            }

        }
        private void hook_Test_DiodeOpen() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Diode_Open_Test\n");
                    receive_present_state = ReportReceive_states.Req_Diode;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("Diode_Open_Test\n");

                    read_result_Diode(JsonSerializer.Deserialize<Diode_Test>(receive_Bluetooth()));
                    break;
            }


        }

        private void hook_Test_AB_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_A_to_B\n");
                    receive_present_state = ReportReceive_states.Req_AB;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    receive_present_state = ReportReceive_states.Req_AB;
                    updateProgress("Testing State A to B...", 100 / 7);
                    if (scan_read) {
                        send_Bluetooth("State_A_to_B\n");

                    } else {
                        send_Bluetooth("State_A_to_B_Single\n");
                    }
                    //State_Transition_Test state_Transition_Test = JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth());
                    ////read_result_A_to_B(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));
                    //read_result_A_to_B(state_Transition_Test);
                    ////CP_Sampling
                    //send_Bluetooth("CP_Sampling");
                    //String jsonArray = receive_Array_Bluetooth();
                    ////float[] cp_array_sample = JsonSerializer.Deserialize<float[]>(jsonArray);
                    //cp_sample_AB = JsonSerializer.Deserialize<float[]>(jsonArray);
                    ////Test_CP_Sample.Text = cp_array_sample.Length.ToString();
                    ////Waveform_pic
                    ////CP_B2_Pic.Image = DrawOscilloscope(cp_sample_AB, CP_B2_Pic.Width, CP_B2_Pic.Height, 50000.0f);
                    //AB_Time.Checked = true;

                    break;
            }
        }
        private void hook_Test_BC_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_B_to_C\n");
                    receive_present_state = ReportReceive_states.Req_BC;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    receive_present_state = ReportReceive_states.Req_BC;
                    updateProgress("Testing State B to C...", 300 / 7);
                    if (scan_read) {
                        send_Bluetooth("State_B_to_C\n");
                    } else {
                        send_Bluetooth("State_B_to_C_Single\n");
                    }

                    //read_result_B_to_C(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));
                    ////CP_Sampling
                    //send_Bluetooth("CP_Sampling");
                    //String jsonArray = receive_Array_Bluetooth();
                    ////float[] cp_array_sample = JsonSerializer.Deserialize<float[]>(jsonArray);
                    //cp_sample_BC = JsonSerializer.Deserialize<float[]>(jsonArray);
                    ////Test_CP_Sample.Text = cp_array_sample.Length.ToString();
                    ////Waveform_pic
                    //BC_Time.Checked = true;
                    ////CP_C2_Pic.Image = DrawOscilloscope(cp_sample_BC, CP_C2_Pic.Width, CP_C2_Pic.Height, 50000.0f);


                    break;
            }

        }
        private void hook_Test_CB_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_C_to_B\n");
                    receive_present_state = ReportReceive_states.Req_CB;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    receive_present_state = ReportReceive_states.Req_CB;
                    if (scan_read) {
                        send_Bluetooth("State_C_to_B\n");
                    } else {
                        send_Bluetooth("State_C_to_B_Single\n");
                    }

                    //read_result_C_to_B(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));

                    //Enable_All_Test_BTN(true);
                    //scan_read = false;
                    break;
            }

        }
        private void hook_Test_BD_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_B_to_D\n");
                    receive_present_state = ReportReceive_states.Req_BD;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    receive_present_state = ReportReceive_states.Req_BD;
                    updateProgress("Testing State C to D...", 500 / 7);
                    if (scan_read) {
                        send_Bluetooth("State_B_to_D\n");
                    } else {
                        send_Bluetooth("State_B_to_D_Single\n");
                    }

                    //read_result_B_to_D(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));
                    ////CP_Sampling
                    //send_Bluetooth("CP_Sampling");
                    //String jsonArray = receive_Array_Bluetooth();
                    ////float[] cp_array_sample = JsonSerializer.Deserialize<float[]>(jsonArray);
                    //cp_sample_CD = JsonSerializer.Deserialize<float[]>(jsonArray);
                    ////Test_CP_Sample.Text = cp_array_sample.Length.ToString();
                    ////Waveform_pic
                    ////CP_D_Pic.Image = DrawOscilloscope(cp_sample_CD, CP_D_Pic.Width, CP_D_Pic.Height, 50000.0f);
                    //CD_Time.Checked = true;


                    break;
            }

        }
        private void hook_Test_RCD_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("RCD0_Test\n");
                    receive_present_state = ReportReceive_states.Req_RCD;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    receive_present_state = ReportReceive_states.Req_RCD;
                    send_Bluetooth("RCD0_Test\n");

                    //read_result_RCD(JsonSerializer.Deserialize<RCD0>(receive_Bluetooth()));
                    break;
            }

        }
        private void hook_Test_Insulator_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Insulator_Test\n");
                    receive_present_state = ReportReceive_states.Req_Insul;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("Insulator_Test\n");

                    //read_result_Insulator(JsonSerializer.Deserialize<Insulation_Test>(receive_Bluetooth()));
                    break;
            }

        }
        private void hook_Test_LinePE_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Test_LinePE\n");
                    receive_present_state = ReportReceive_states.Req_Insul;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("Test_LinePE\n");

                    //read_result_LinePE(JsonSerializer.Deserialize<Insulation_Test>(receive_Bluetooth()));
                    break;
            }
        }
        private void hook_Test_NeutPE_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Test_NeutralPE\n");
                    receive_present_state = ReportReceive_states.Req_Insul;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    send_Bluetooth("Test_NeutralPE\n");

                    //read_result_NeutPE(JsonSerializer.Deserialize<Insulation_Test>(receive_Bluetooth()));
                    break;
            }
        }
        private void hook_Test_Diode_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Diode_Test\n");
                    receive_present_state = ReportReceive_states.Req_Diode;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    receive_present_state = ReportReceive_states.Req_Diode;
                    send_Bluetooth("Diode_Test\n");

                    //read_result_Diode(JsonSerializer.Deserialize<Diode_Test>(receive_Bluetooth()));
                    break;
            }

        }
        private void hook_Test_PE_Open_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("PE_Open_Test\n");
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    receive_present_state = ReportReceive_states.Req_Diode;
                    send_Bluetooth("PE_Open_Test\n");

                    //read_result_Diode(JsonSerializer.Deserialize<Diode_Test>(receive_Bluetooth()));
                    break;
            }

        }
        private void hook_Test_DiodeOpen_nonBlocking() {
            switch (commu_mode) {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("Diode_Open_Test\n");
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    receive_present_state = ReportReceive_states.Req_Diode;
                    send_Bluetooth("Diode_Open_Test\n");

                    //read_result_Diode(JsonSerializer.Deserialize<Diode_Test>(receive_Bluetooth()));
                    break;
            }


        }


        private void update_CP_B2_Pic(Image newImage) {
            if (CP_B2_Pic.InvokeRequired) {
                Action action = delegate {
                    using (Bitmap bmp = new Bitmap(newImage)) {
                        CP_B2_Pic.Image?.Dispose(); // Dispose the old image to free resources
                        CP_B2_Pic.Image = new Bitmap(bmp);
                    }
                };
                CP_B2_Pic.Invoke(action);
            } else {
                using (Bitmap bmp = new Bitmap(newImage)) {
                    CP_B2_Pic.Image?.Dispose(); // Dispose the old image to free resources
                    CP_B2_Pic.Image = new Bitmap(bmp);
                }
            }
        }
        private void update_CP_C2_Pic(Image newImage) {
            if (CP_C2_Pic.InvokeRequired) {
                Action action = delegate {
                    using (Bitmap bmp = new Bitmap(newImage)) {
                        CP_C2_Pic.Image?.Dispose(); // Dispose the old image to free resources
                        CP_C2_Pic.Image = new Bitmap(bmp);
                    }
                };
                CP_C2_Pic.Invoke(action);
            } else {
                using (Bitmap bmp = new Bitmap(newImage)) {
                    CP_C2_Pic.Image?.Dispose(); // Dispose the old image to free resources
                    CP_C2_Pic.Image = new Bitmap(bmp);
                }
            }
        }
        private void update_CP_D_Pic(Image newImage) {
            if (CP_D_Pic.InvokeRequired) {
                Action action = delegate {
                    using (Bitmap bmp = new Bitmap(newImage)) {
                        CP_D_Pic.Image?.Dispose(); // Dispose the old image to free resources
                        CP_D_Pic.Image = new Bitmap(bmp);
                    }
                };
                CP_D_Pic.Invoke(action);
            } else {
                using (Bitmap bmp = new Bitmap(newImage)) {
                    CP_D_Pic.Image?.Dispose(); // Dispose the old image to free resources
                    CP_D_Pic.Image = new Bitmap(bmp);
                }
            }
        }

        private void updateProgress(string progress, int value) {
            if (TESTProcessLBL.InvokeRequired) {
                Action action = delegate {
                    TESTProcessLBL.Text = progress;
                };
                TESTProcessLBL.Invoke(action);
            } else {
                TESTProcessLBL.Text = progress;
            }

            if (TESTProcessBar.InvokeRequired) {
                Action action = delegate {
                    TESTProcessBar.Value = value;
                };
                TESTProcessBar.Invoke(action);
            } else {
                TESTProcessBar.Value = value;
            }
        }

        private void updateTestBTN(bool isTesting) {

            //Test_ALL_BTN
            //Test_diode_Front
            //Test_diode_open_Front
            //Test_PE_open_Front
            //RCD_TestBTN
            //cancelBTN
            if (Test_ALL_BTN.InvokeRequired) {
                Action action = delegate {
                    Test_ALL_BTN.Enabled = !isTesting;
                };
                Test_ALL_BTN.Invoke(action);
            } else {
                Test_ALL_BTN.Enabled = !isTesting;
            }

            if (Test_diode_Front.InvokeRequired) {
                Action action = delegate {
                    Test_diode_Front.Enabled = !isTesting;
                };
                Test_diode_Front.Invoke(action);
            } else {
                Test_diode_Front.Enabled = !isTesting;
            }

            if (Test_diode_open_Front.InvokeRequired) {
                Action action = delegate {
                    Test_diode_open_Front.Enabled = !isTesting;
                };
                Test_diode_open_Front.Invoke(action);
            } else {
                Test_diode_open_Front.Enabled = !isTesting;
            }

            if (Test_PE_open_Front.InvokeRequired) {
                Action action = delegate {
                    Test_PE_open_Front.Enabled = !isTesting;
                };
                Test_PE_open_Front.Invoke(action);
            } else {
                Test_PE_open_Front.Enabled = !isTesting;
            }

            if (RCD_TestBTN.InvokeRequired) {
                Action action = delegate {
                    RCD_TestBTN.Enabled = !isTesting;
                };
                RCD_TestBTN.Invoke(action);
            } else {
                RCD_TestBTN.Enabled = !isTesting;
            }

            if (cancelBTN.InvokeRequired) {
                Action action = delegate {
                    cancelBTN.Enabled = isTesting;
                };
                cancelBTN.Invoke(action);
            } else {
                cancelBTN.Enabled = isTesting;

            }
            if (bluetooth_connecting) {
                if (Bluetooth_Connect.InvokeRequired) {
                    Action action = delegate {
                        Bluetooth_Connect.Enabled = !isTesting;
                    };
                    Bluetooth_Connect.Invoke(action);
                } else {
                    Bluetooth_Connect.Enabled = !isTesting;
                }
            }
            
        }

        void clearFrontDisplay_State() {
            B_PWM_Pk.Text = "-";
            B_PWM_nPk.Text = "-";
            B_PWM_Freq.Text = "-";
            B_PWM_Duty.Text = "-";
            B_PWM_Imax.Text = "-";

            B_PWM_Imax_sts.Text = "-";
            B_PWM_Pk_sts.Text = "-";
            B_PWM_nPk_sts.Text = "-";
            B_PWM_Freq_sts.Text = "-";
            B_PWM_Duty_sts.Text = "-";

            B_PWM_Imax_sts.ForeColor = Color.Black;
            B_PWM_Pk_sts.ForeColor = Color.Black;
            B_PWM_nPk_sts.ForeColor = Color.Black;
            B_PWM_Freq_sts.ForeColor = Color.Black;
            B_PWM_Duty_sts.ForeColor = Color.Black;

            C_PWM_Pk.Text = "-";
            C_PWM_nPk.Text = "-";
            C_PWM_Freq.Text = "-";
            C_PWM_Duty.Text = "-";
            C_PWM_Imax.Text = "-";
            C_PWM_MainVolt.Text = "-";
            C_PWM_MainFreq.Text = "-";
            C_PWM_PP.Text = "-";

            C_PWM_Pk_sts.Text = "-";
            C_PWM_nPk_sts.Text = "-";
            C_PWM_Freq_sts.Text = "-";
            C_PWM_Duty_sts.Text = "-";
            C_PWM_Imax_sts.Text = "-";
            C_PWM_MainVolt_sts.Text = "-";
            C_PWM_MainFreq_sts.Text = "-";
            C_PWM_PP_sts.Text = "-";

            C_PWM_Pk_sts.ForeColor = Color.Black;
            C_PWM_nPk_sts.ForeColor = Color.Black;
            C_PWM_Freq_sts.ForeColor = Color.Black;
            C_PWM_Duty_sts.ForeColor = Color.Black;
            C_PWM_Imax_sts.ForeColor = Color.Black;
            C_PWM_MainVolt_sts.ForeColor = Color.Black;
            C_PWM_MainFreq_sts.ForeColor = Color.Black;
            C_PWM_PP_sts.ForeColor = Color.Black;

            D_PWM_Pk.Text = "-";

            D_PWM_nPk.Text = "-";
            D_PWM_Freq.Text = "-";
            D_PWM_Duty.Text = "-";
            D_PWM_Imax.Text = "-";

            D_PWM_MainVolt.Text = "-";
            D_PWM_MainFreq.Text = "-";
            D_PWM_PP.Text = "-";
            D_PWM_Pk_sts.Text =  "-";
            D_PWM_nPk_sts.Text = "-";
            D_PWM_Freq_sts.Text = "-";
            D_PWM_Duty_sts.Text = "-";
            D_PWM_Imax_sts.Text = "-";
            D_PWM_MainVolt_sts.Text = "-";
            D_PWM_MainFreq_sts.Text = "-";
            D_PWM_PP_sts.Text = "-";
            D_PWM_Pk_sts.ForeColor = Color.Black;
            D_PWM_nPk_sts.ForeColor = Color.Black;
            D_PWM_Freq_sts.ForeColor = Color.Black;
            D_PWM_Duty_sts.ForeColor = Color.Black;
            D_PWM_Imax_sts.ForeColor = Color.Black;
            D_PWM_MainVolt_sts.ForeColor = Color.Black;
            D_PWM_MainFreq_sts.ForeColor = Color.Black;
            D_PWM_PP_sts.ForeColor = Color.Black;

            CB_PWM_MainOff.Text = "-";
            CB_PWM_MainOff_sts.Text = "-";
            CB_PWM_MainOff_sts.ForeColor = Color.Black;

            BC_MainOn.Text = "-";
            BC_MainOn_sts.Text = "-";
            BC_MainOn_sts.ForeColor = Color.Black;
        }
        void clearFrontDisplay_Diode_Short() {
            DiodeSh_MainOff.Text = "-";
            DiodeSh_MainOff_sts.Text = "-";
            DiodeSh_MainOff_sts.ForeColor = Color.Black;
        }
        void clearFrontDisplay_PE_Open() {
            PE_Open_MainOff.Text = "-";
            PE_Open_MainOff_sts.Text = "-";
            PE_Open_MainOff_sts.ForeColor = Color.Black;
        }
        void clearFrontDisplay_Diode_Open() {
            DiodeOpen_MainOff.Text = "-";
            DiodeOpen_MainOff_sts.Text = "-";
            DiodeOpen_MainOff_sts.ForeColor = Color.Black;
        }
        void clearFrontDisplay_RCD() {
            RCD_TripTime_ms.Text = "-";
            RCD_Accecptable_TripTime.Text = "-";
            RCD_TestingInjectedCurrent.Text = "-";

            RCD_TripTime_ms_sts.Text = "-";
            RCD_TripTime_ms_sts.ForeColor = Color.Black;
        }
    }
}
