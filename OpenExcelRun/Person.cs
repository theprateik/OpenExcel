﻿using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun
{
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public double Income { get; set; }
        public DateTime DateOfBirth { get; set; }
        public List<Child> Children { get; set; }
    }
}
