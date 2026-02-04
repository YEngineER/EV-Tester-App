using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;

namespace EV_Testing_report_DEMO
{
    public partial class Form1 : Form
    {

        private void hook_Test_AB()
        {
            switch (commu_mode)
            {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_A_to_B\n");
                    receive_present_state = ReportReceive_states.Req_AB;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    if (scan_read)
                    {
                        send_Bluetooth("State_A_to_B\n");
                    }
                    else
                    {
                        send_Bluetooth("State_A_to_B_Single\n");
                    }
                    read_result_A_to_B(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));
                    //CP_Sampling
                    send_Bluetooth("CP_Sampling");
                    String jsonArray = receive_Array_Bluetooth();
                    //float[] cp_array_sample = JsonSerializer.Deserialize<float[]>(jsonArray);
                    cp_sample_AB = JsonSerializer.Deserialize<float[]>(jsonArray);
                    //Test_CP_Sample.Text = cp_array_sample.Length.ToString();
                    //Waveform_pic
                    //CP_B2_Pic.Image = DrawOscilloscope(cp_sample_AB, CP_B2_Pic.Width, CP_B2_Pic.Height, 50000.0f);
                    AB_Time.Checked = true;

                    if (scan_read)
                    {
                        hook_Test_BC();
                    }

                    break;
            }
        }
        private void hook_Test_BC()
        {
            switch (commu_mode)
            {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_B_to_C\n");
                    receive_present_state = ReportReceive_states.Req_BC;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    if (scan_read)
                    {
                        send_Bluetooth("State_B_to_C\n");
                    }
                    else
                    {
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
                    if (scan_read)
                    {
                        hook_Test_BD();
                    }

                    break;
            }

        }
        private void hook_Test_CB()
        {
            switch (commu_mode)
            {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_C_to_B\n");
                    receive_present_state = ReportReceive_states.Req_CB;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    if (scan_read)
                    {
                        send_Bluetooth("State_C_to_B\n");
                    }
                    else
                    {
                        send_Bluetooth("State_C_to_B_Single\n");
                    }

                    read_result_C_to_B(JsonSerializer.Deserialize<State_Transition_Test>(receive_Bluetooth()));

                    Enable_All_Test_BTN(true);
                    scan_read = false;
                    break;
            }

        }
        private void hook_Test_BD()
        {
            switch (commu_mode)
            {
                case EVSE_Tester_CommunicationMode.None:
                    break;
                case EVSE_Tester_CommunicationMode.SerialPort:
                    esp32_module.WriteLine("State_B_to_D\n");
                    receive_present_state = ReportReceive_states.Req_BD;
                    break;
                case EVSE_Tester_CommunicationMode.Bluetooth:
                    if (scan_read)
                    {
                        send_Bluetooth("State_B_to_D\n");
                    }
                    else
                    {
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
                    if (scan_read)
                    {
                        hook_Test_CB();
                    }

                    break;
            }

        }
        private void hook_Test_RCD()
        {
            switch (commu_mode)
            {
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
        private void hook_Test_Insulator()
        {
            switch (commu_mode)
            {
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
        private void hook_Test_LinePE()
        {
            switch (commu_mode)
            {
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
        private void hook_Test_NeutPE()
        {
            switch (commu_mode)
            {
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
        private void hook_Test_Diode()
        {
            switch (commu_mode)
            {
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

        private void hook_Test_PE_Open()
        {
            switch (commu_mode)
            {
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

        private void hook_Test_DiodeOpen()
        {
            switch (commu_mode)
            {
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

    }
}
