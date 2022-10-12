using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    internal class Validation
    {
        /// <summary>
        /// Ensure that all fields have a non-empty value
        /// </summary>
        /// <param name="listItem"></param>
        /// <returns>String list of validation errors, otherwise empty string</returns>
        public static string ValidateInput(FieldValueSet fieldValueSet)
        {
            string ValidationErrors = String.Empty;
            foreach (string k in fieldValueSet.AdditionalData.Keys)
            {
                if ((fieldValueSet.AdditionalData[k] is null) || string.IsNullOrEmpty(fieldValueSet.AdditionalData[k].ToString().Trim()))
                {
                    ValidationErrors += string.Format("Field {0} cannot be blank.", k) + Environment.NewLine;
                }
            }

            return ValidationErrors;
        }

    }
}
