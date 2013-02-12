/*
 * Use WCF and LINQ to Sharepoint to integrate sharepoint
 * Author: Alessandro Graps
 * Year: 2013
 */using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace sharepointIntegration
{
    public static class CommonClass
    {
        /// <summary>
        /// Prints the properties.
        /// </summary>
        /// <param name="obj">The obj.</param>
        public static void PrintProperties(object obj)
        {
            PrintProperties(obj, 0);
        }

        /// <summary>
        /// Prints the properties.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="indent">The indent.</param>
        public static void PrintProperties(object obj, int indent)
        {
            if (obj == null) return;
            string indentString = new string(' ', indent);
            Type objType = obj.GetType();
            PropertyInfo[] properties = objType.GetProperties();
            foreach (PropertyInfo property in properties)
            {
                object propValue = property.GetValue(obj, null);
                if (property.PropertyType.Assembly == objType.Assembly)
                {
                    Console.WriteLine("{0}{1}:", indentString, property.Name);
                    PrintProperties(propValue, indent + 2);
                }
                else
                {
                    Console.WriteLine("{0}{1}: {2}", indentString, property.Name, propValue);
                }
            }
        }
            
        
    }
}
