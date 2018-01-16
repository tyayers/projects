﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DocBot.Dtos
{
    public class QnARequest
    {
        public string question { get; set; }
    }
    public class QnAResponse
    {
        public QnAAnswer[] answers { get; set; }
    }
    public class QnAAnswer
    {
        public string answer { get; set; }
        public decimal score { get; set; }
    }
}