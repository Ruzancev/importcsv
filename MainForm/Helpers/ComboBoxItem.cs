﻿using Microsoft.SharePoint.Client;


namespace Trilogen.Helpers
{
    public class ComboBoxItem
    {
        public string Text { get; set; }
        public object Value { get; set; }
        public FieldCollection Fields { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }
}
