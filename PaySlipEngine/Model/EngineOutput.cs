﻿namespace PaySlipEngine.Model
{
    public class EngineOutput
    {
        public string Name { get; set; }

        public decimal GrossIncome { get; set; }   

        public decimal IncomeTax { get; set; }

        public decimal NetIncome { get; set; }

        public decimal Super { get; set; }

        public string PayPeriod { get; set; }
    }
}
