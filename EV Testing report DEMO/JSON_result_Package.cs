using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;

namespace EV_Testing_report_DEMO
{
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
    
}
