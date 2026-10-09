using System;

namespace PJShoesSlipRecords
{
    public sealed class SlipRecord
    {
        public int SlipNo { get; set; }
        public string CustomerName { get; set; }
        public string PhoneNo { get; set; }
        public DateTime EntryDate { get; set; }
        public DateTime CreatedAt { get; set; }

        public SlipRecord Copy()
        {
            return new SlipRecord
            {
                SlipNo = SlipNo,
                CustomerName = CustomerName,
                PhoneNo = PhoneNo,
                EntryDate = EntryDate,
                CreatedAt = CreatedAt
            };
        }
    }
}
