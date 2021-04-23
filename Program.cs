
/*
LICENSE INFORMATION

Copyright 2020 Kaizen Cyber Ops, LLC.

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

https://opensource.org/licenses/MIT
*/


using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using OutlookClient.Structures;

namespace OutlookClient
{
    public enum SendAs { FILE_ATTACHMENT, PLAINTEXT_BODY }

    class OutlookClient
    {

        private string email_body;
        private readonly string subject_line;
        public OutlookClient(string email_body, string subject_line)
        {
            this.email_body = email_body;
            this.subject_line = subject_line;
        }



        private void SendMailItem(string email_to, string full_body)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = subject_line;
            mailItem.To = email_to;
            mailItem.Body = full_body;
            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            mailItem.Display(false);
            mailItem.Send();
        }

        public int SendAllRecords(DataRecordCollection data, string destination)
        {
            string sending_body = email_body;
            sending_body += Environment.NewLine;
            sending_body += data.ToString();
            SendMailItem(destination, sending_body);
            return data.GetShape().Rows;
        }

        public int SendDecrementalEmails(DataRecordCollection data, string destination)
        {
            int sent = 0;
            int current_chunk = data.GetShape().Rows;
            while (sent < data.GetShape().Rows)
            {
                current_chunk /= 2;
                if (current_chunk == 0)
                    ++current_chunk;    //we need at least one more record if its here and the chunk size is zero
                int records_to_send = 0;
                if (sent + current_chunk >= data.GetShape().Rows)
                    records_to_send = data.GetShape().Rows - sent;
                else
                    records_to_send = current_chunk;
                DataRecordCollection drc = new DataRecordCollection(data.Schema);
                foreach (DataRecord d in data.GetRecords().Skip(sent).Take(records_to_send))
                    drc.AddRecord(d);
                SendAllRecords(drc, destination);
                sent += records_to_send;
            }
            return sent;
        }

        public int Send(DataRecordCollection data, string destination)
        {
            int sent = 0;
            foreach (DataRecord d in data.GetRecords())
            {
                string sending_body = email_body + d.ToString();
                SendMailItem(destination, sending_body);
                sent++;
            }
            return sent;
        }
    }


}

namespace OutlookClient.Structures
{
    /// <summary>
    /// This is what holds multiple DataRecord objects. 
    /// It is treated as an enumerable 2 dimensional array which enumerates DataRecords
    /// The column size and row space of the DataRecordCollection is defined by its Shape()
    /// </summary>
    public class DataRecordCollection : IEnumerable
    {
        private string[] schema;
        private DataShape shape;
        private DataRecord[] records;
        private int Columns { get => shape.Columns; }
        private int Rows { get => shape.Rows; }
        public string[] Schema { get => schema; }
        public DataRecordCollection(string[] scheme)
        {
            schema = scheme;
            shape = new DataShape(schema.Length, 0);
            records = new DataRecord[100];
        }
        public DataRecordCollection(string[] scheme, string[][] recs)
        {
            schema = scheme;
            shape = new DataShape(schema.Length, recs.Length);
            records = new DataRecord[100];
            foreach (String[] s in recs)
                this.AddRecord(s);
        }
        public DataRecordCollection(string[] scheme, string[] rec)
        {
            schema = scheme;

            shape = new DataShape(schema.Length, 1);
            records = new DataRecord[100];
            AddRecord(rec);
        }
        public void AddRecord(DataRecord rec)
        {
            if (rec.Length != shape.Columns)
                throw new ArgumentException(String.Format("The input record has {0} fields which is incompatible with the current schema size of {1} fields", rec.Length, shape.Columns));
            while (records.Length < shape.Rows + 1)
                Array.Resize<DataRecord>(ref records, records.Length + 100);
            records[shape.Rows] = rec;
            shape.Increment();
        }
        public void AddRecord(string[] rec)
        {
            AddRecord(new DataRecord(rec));
        }
        public void AddAllRecords(string[][] rec)
        { foreach (string[] r in rec) AddRecord(r); }
        public void AddAllRecords(DataRecord[] recs)
        { foreach (DataRecord r in recs) AddRecord(r); }
        public IEnumerator<DataRecord> GetDataRecordEnumerator()
        { return (IEnumerator<DataRecord>)records.GetEnumerator(); }
        public IEnumerator GetEnumerator()
        { return records.GetEnumerator(); }
        public IEnumerator GetColumnEnumerator()
        { return schema.GetEnumerator(); }
        public IEnumerator GetDictEnumerator()
        {
            Dictionary<string[], DataRecord> AsDict = new Dictionary<string[], DataRecord>();
            for (int i = 0; i < Rows; i++)
                AsDict.Add(schema, records[i]);
            return AsDict.GetEnumerator();
        }
        public DataRecord[] GetRecords() { return records; }

        public DataShape GetShape() { return shape; }
        public string[] GetSchema() { return schema; }
        override public string ToString()
        {
            return (schema != null && records.Length > 0) //is there a schema and records?  if so, turn it into a csv-style format.  else, return an empty string.
                ? String.Join(", ", schema) + Environment.NewLine + records.Select(r => r.ToString() + Environment.NewLine)
                : "";
        }

    }


    /// <summary>
    /// This object represents an individual DataRecord, which is an abstraction of the data within a file.
    /// The data it represents must be structured within the file such that it can be represented as a collection of key-value pairs
    /// Any data that can be represented in 2-dimensional or multidimensional tables such as CSV, SQL, JSON, etc... are ideal.
    /// Unstructured data may be represented by simply assigning every section of unstructured data an arbitrary (or the same) key, but this is non-ideal
    /// </summary>
    public class DataRecord : IEnumerable
    {
        public string[] Data { get; private set; }
        public readonly int size;
        public int Length { get => size; }
        public DataRecord(string[] values)
        {
            size = values.Length;
            Data = values;
        }
        public IEnumerator GetEnumerator() { return Data.GetEnumerator(); }
        public IEnumerator<string> GetFieldEnumerator() { return (IEnumerator<string>)Data.GetEnumerator(); }
        override public string ToString()
        { return String.Join(", ", Data); }
    }


    /// <summary>
    /// This descriptes the shape of a DataRecordCollection.
    /// It models a DataRecordCollection as a 2d array and gives you the size of the array as (Rows,Cols)
    /// </summary>
    public class DataShape
    {
        public int Columns { get; private set; }
        public int Rows { get; private set; }
        public DataShape(int item1, int item2)
        {
            Columns = item1;
            Rows = item2;
        }
        public void Increment()
        { ++Rows; }
    }




    /// <summary>
    /// class Schema is what represents the schema of the data in question.  Obviously this will fall apart if trying to handle unstructured data.
    /// </summary>
    public class Schema
    {

        public string[] schema { get; private set; }

        /// <summary>
        /// Constructor that will split a single string into an array as per a delimiter.  the invidual array members will be the schema.
        /// </summary>
        /// <param name="str">represents the string to be split</param>
        /// <param name="delimiter">represents the delimiter to split by</param>
        /// 
        public Schema(string str, char delimiter)
        { schema = str.Split(delimiter); }


        /// <summary>
        /// Takes a strig array and assigns that as the schema.  no additional work will be done when assigning the schema.
        /// </summary>
        /// <param name="str"> the string array to use as the schema</param>
        /// 
        public Schema(string[] str)
        { schema = str; }


        /// <summary>
        /// Takes a string and splits according to a regex pattern into an array, using Regex.Split() method
        /// </summary>
        /// <param name="str">the string to be split</param>
        /// <param name="split_pattern">the Regex pattern to use for splitting</param>
        /// 
        public Schema(string str, Regex split_pattern)
        { schema = Regex.Split(str, split_pattern.ToString()); }


        /// <summary>
        /// Takes a string and assigns the schema according to the result of splitting str using split_string
        /// </summary>
        /// <param name="str">the string to be split to a schema</param>
        /// <param name="split_string">the string to be used to split str</param>
        public Schema(string str, string split_string)
        { schema = str.Split(new string[] { split_string }, StringSplitOptions.None); }

    }

}

