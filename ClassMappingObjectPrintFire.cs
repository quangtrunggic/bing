using Aspose.Words;
using Newtonsoft.Json;
using PagedList;
using PRP_SCCM.Models;
using PRP_SCCM.Utilities;
using PRP_SCCM.ViewModels;
using PRP_SCCM.ViewModels.PrintPolicy;
using PRP_SCCM.ViewModels.PrintPolicy.ObjectFire;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;

namespace PRP_SCCM.Classes
{

    public class ClassMappingObjectPrintFire
    {
        #region Variables
        private CommonTemplateGicore _commonTemplateGicore = new CommonTemplateGicore();
        private Root _root = new Root();
        private DataModelView _dataModelView = new DataModelView();
        #endregion
        // Hàm trả về idtemplate mẫu in theo Ký hiệu Data (UI)
        private void GetIdTemplate(List<GICORE_METADATA> lstCode, Guid idtemplateParent)
        {
            var idtem = Guid.NewGuid();
            if (lstCode.Where(x => x.Code.Contains("GR")).Any())
            {
                //if (_root.CERTIFICATES.CERTIFICATE.Count == 1)
                //{
                //    idtem = Guid.Parse(_commonTemplateGicore.GetIDTemplate("GRO", idtemplateParent));
                //    _dataModelView.idtemplate = idtem;
                //}
                //else
                //{
                //    idtem = Guid.Parse(_commonTemplateGicore.GetIDTemplate("GRM", idtemplateParent));
                //    _dataModelView.idtemplate = idtem;
                //}
            }
            else
            {
                idtem = Guid.Parse(_commonTemplateGicore.GetIDTemplate("BT", idtemplateParent));
                _dataModelView.idtemplate = idtem;
            }
        }
        // Hàm trả về Root Object chứa data XML file
        private void GenerateRoot(string filename)
        {
            string localDirectory = ConfigurationManager.AppSettings["pathFileXML"];
            XmlSerializer serializer = new XmlSerializer(typeof(Root));
            _root = (Root)serializer.Deserialize(new XmlTextReader(localDirectory + @"\" + filename));
        }
        string GetAddress(List<ADDRESS> contact)
        {
            string address = "";
            foreach (var item in contact)
            {
                if (item.IsPrimaryAddress == "Y" || item.IsPrimaryAddress == null)
                {
                    if (item.Address != null)
                        address += item.Address + ", ";
                    if (item.Ward != null)
                        address += item.Ward + ", ";
                    if (item.District != null)
                        address += item.District + ", ";
                    if (item.City != null)
                        address += item.City + ", ";
                }
            }
            return address.TrimEnd(',', ' ');
        }
        // Hàm trả về Content Object hiển thị trên Word
        public DataModelView MappingDataObjectPrint(string filename, string dataSymbol, Guid idtemplateParent, string language)
        {
            // Nạp _root
            GenerateRoot(filename);
            #region Variables
            var certificate = _root.CERTIFICATES.CERTIFICATE;
            var certificates = _root.CERTIFICATES;
            var groupPolicy = _root.GROUPPOLICY;
            string productCode = _root.GROUPPOLICY.ProductCode,
                   currencyCode = _root.GROUPPOLICY.CurrencyCode;
            Boolean hasCOVFCI, // Cháy Nổ Bắt Buộc
                    hasCOVFSP, // Bảo hiểm hỏa hạn và các rủi ro đặc biệt
                    hasCOVFPA, // Mọi Rủi Ro Tài Sản
                    hasCOVFBI; // Gián đoạn kinh doanh
            #endregion
            // Nạp idtemplate
            GetIdTemplate(_commonTemplateGicore.GetListCode(idtemplateParent), idtemplateParent);
            #region object Excel
            //if (!string.IsNullOrEmpty(dataSymbol))
            //{
            //    _dataModelView.dt_tableExcel = ExportDataExcel(_root, dataSymbol, lstCode, idtemplateParent, ref idtem);
            //    _dataModelView.idtemplate = idtem;
            //    return _dataModelView;
            //}
            #endregion


            PolicyInformation Policy()
            {
                // Khai báo các biến
                var policy = new Dictionary<string, string>();

                // Lấy thông tin các điều khoản bảo hiểm
                GetClauseInfo(policy);

                // Lấy thông tin ngân hàng của GIC
                GetGICBankInfo(policy);

                // Lấy thông tin khách hàng
                GetCustomerInfo(policy);

                // Lấy thông tin rủi ro bảo hiểm
                GetRiskInfo(policy);

                // Lấy thông tin phí bảo hiểm
                GetFeeInfo(policy);

                // Lấy thông tin phí bảo hiểm giấy chứng nhận
                GetCertificateFeeInfo(policy);

                // Tạo đối tượng PolicyInformation từ dictionary policy
                PolicyInformation policyInformation = CreatePolicyInformation(policy);

                return policyInformation;
            }
            _dataModelView.pros = DataDictionaryExtension.BuildMergedFields(Policy());
            
            // Hàm lấy thông tin các điều khoản bảo hiểm
            void GetClauseInfo(Dictionary<string, string> policy)
            {
                #region Clause: is_Clause, is_ClauseDKBS, is_ClauseNTH, stt, is_Bi
                string is_Clause = "N",
                       is_ClauseDKBS = "N",
                       is_ClauseNTH = "N",
                       stt4 = "2",
                       stt5 = "3",
                       stt6 = "4",
                       stt7 = "7",
                       stt8 = "8",
                       stt9 = "9",
                       is_Bi = "N";
                #region is_Clause 
                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if (string.IsNullOrEmpty(item.ClauseCoverageCodes) && string.IsNullOrEmpty(item.PrintOrder))
                    {
                    }
                    else if (string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && item.ClauseCoverageCodes == "CCHD")
                    {
                        is_Clause = "Y";
                    }
                    else if (string.IsNullOrEmpty(item.PrintOrder) == false && item.PrintOrder == "CCHD")
                    {
                        is_Clause = "Y";
                    }
                }
                #endregion
                #region is_ClauseDKBS
                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if (string.IsNullOrEmpty(item.ClauseContent) == false)
                    {
                        is_ClauseDKBS = "Y";
                    }
                    else if (string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && item.ClauseCoverageCodes == "CCHD")
                    {
                        is_ClauseDKBS = "Y";
                    }
                    else if (string.IsNullOrEmpty(item.PrintOrder) == false && item.PrintOrder == "CCHD")
                    {
                        is_ClauseDKBS = "Y";
                    }
                }
                #endregion
                #region is_ClauseNTH, _stt
                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if ((string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && item.ClauseCoverageCodes == "NTH")
                        || (string.IsNullOrEmpty(item.PrintOrder) == false && item.PrintOrder == "NTH")
                        )
                    {
                        is_ClauseNTH = "Y";
                        stt4 = "3";
                        stt5 = "4";
                        stt6 = "5";
                        stt7 = "8";
                        stt8 = "9";
                        stt9 = "10";
                    }
                }
                #endregion
                #region _is_Bi
                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if (string.IsNullOrEmpty(item.ClauseCoverageCodes) && string.IsNullOrEmpty(item.PrintOrder))
                    {
                    }
                    else if (string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && (item.ClauseCoverageCodes.Contains("BI") || item.ClauseCoverageCodes.Contains("BI")))
                    {
                        is_Bi = "Y";
                    }
                    else if (string.IsNullOrEmpty(item.PrintOrder) == false && (item.PrintOrder.Contains("BI") || item.PrintOrder.Contains("BI")))
                    {
                        is_Bi = "Y";
                    }
                }
                #endregion
                policy.Add("IsClause", is_Clause);
                policy.Add("IsClauseDKBS", is_ClauseDKBS);
                policy.Add("IsClauseNTH", is_ClauseNTH);
                policy.Add("stt4", stt4);
                policy.Add("stt5", stt5);
                policy.Add("stt6", stt6);
                policy.Add("stt7", stt7);
                policy.Add("stt8", stt8);
                policy.Add("stt9", stt9);
                policy.Add("IsBi", is_Bi);
                #endregion
            }
             
            // Hàm lấy thông tin ngân hàng của GIC
            void GetGICBankInfo(Dictionary<string, string> policy)
            {
                string gicBranchBankAccount,
                       isGICBranchBankAccount,
                       beneficiary_Name = "",
                       beneficiary_Detail = "",
                       Beneficiary_Detail2 = "";

                string[] separatingStrings = { "<br/>" };
                gicBranchBankAccount = groupPolicy.GICBranchBankAccount ?? "";
                isGICBranchBankAccount = (groupPolicy.GICBranchBankAccount != null && (("40274057, 22222999702, 38783878888, 0100100035953002, 059704070427888, 100003315999999, 210314851015105").Contains(gicBranchBankAccount))) ? "Y" : "N";

                if (groupPolicy.Beneficiary != null)
                {
                    var beneficiaryParts = groupPolicy.Beneficiary.Split(separatingStrings, System.StringSplitOptions.RemoveEmptyEntries);
                    beneficiary_Name = beneficiaryParts[0];
                    beneficiary_Detail = string.Join("                                                                                                                                                                    ", beneficiaryParts.Skip(1)).Replace("<br/>", "                                                                                                                                                                    ");
                    Beneficiary_Detail2 = groupPolicy.Beneficiary != null ? groupPolicy.Beneficiary.Split(separatingStrings, System.StringSplitOptions.RemoveEmptyEntries)[1].Replace("<br/>", "                                                                                                                                                                    ") : "";
                }

                policy.Add("GICBranchBankAccount", gicBranchBankAccount);
                policy.Add("IsGICBranchBankAccount", isGICBranchBankAccount);
                policy.Add("Beneficiary_Name", beneficiary_Name);
                policy.Add("Beneficiary_Detail", beneficiary_Detail);
                policy.Add("Beneficiary_Detail2", Beneficiary_Detail2);
            }

            // Hàm lấy thông tin khách hàng
            void GetCustomerInfo(Dictionary<string, string> policy)
            {
                string customer_PartyNameVN = "",
                       customer_Address = "",
                       customer_PartyNameVN_S = "",
                       customer_Address_S = "",
                       customer_Phone = "",
                       customer_Phone_S = "",
                       customerFax = "",
                       customerFax_S = "",
                       customerIDNumber = "",
                       customerIDNumber_S = "",
                       customerRegistrationNumber = "",
                       customerRegistrationNumber_S = "",
                       customerVATRegistrationNumber = "",
                       customerVATRegistrationNumber_S = "",
                       customerRepresen = "",
                       customerRepresen_S = "",
                       customerPosition = "",
                       customerPosition_S = "",
                       customer_AttorneyNo = "",
                       customer_AttorneyNo_S = "",
                       isInsuredAndPolicyHolder = "N",
                       isIndividual = "Y",
                       isIndividual_S = "Y";

                foreach (var item in groupPolicy.CUSTOTMER)
                {
                    bool hasINDIVIDUAL = item.INDIVIDUAL_CUSTOMER != null;
                    bool hasORGANIZATION = item.ORGANIZATION_CUSTOMER != null;
                    bool hasREPRESENTATIVE = hasORGANIZATION && item.ORGANIZATION_CUSTOMER.REPRESENTATIVE != null;
                    bool hasInsured = item.IsInsured != null;
                    bool hasPolicyHolder = item.IsPolicyHolder != null;

                    // Kiểm tra người mua bảo hiểm và người được bảo hiểm có trùng nhau không
                    if (hasInsured && hasPolicyHolder && item.IsInsured == "Y" && item.IsPolicyHolder == "Y")
                    {
                        isInsuredAndPolicyHolder = "Y";
                    }

                    // Lấy thông tin người mua bảo hiểm
                    if (item.IsPolicyHolder == "Y")
                    {
                        string isCustomer_Phone = "N",
                            isCustomer_IDNumber = "N",
                            isCustomer_RegistrationNumber = "N",
                            isCustomer_VATRegistrationNumber = "N",
                            isCustomer_Represen = "N",
                            isCustomer_Position = "N",
                            isCustomer_AttorneyNo = "N";

                        // Kiểm tra người được bảo hiểm là cá nhân hay tổ chức
                        if (hasORGANIZATION)
                        {
                            isIndividual = "N";
                        }
                        else if (hasINDIVIDUAL)
                        {
                            isIndividual = "Y";
                        }

                        // Lấy tên khách hàng
                        if (hasINDIVIDUAL)
                        {
                            customer_PartyNameVN = item.INDIVIDUAL_CUSTOMER.PartyName;
                        }
                        else if (hasORGANIZATION)
                        {
                            customer_PartyNameVN = item.ORGANIZATION_CUSTOMER.PartyNameVN;
                        }

                        // Lấy địa chỉ khách hàng
                        if (hasINDIVIDUAL)
                        {
                            foreach (var address in item.INDIVIDUAL_CUSTOMER.ADDRESS)
                            {
                                if (address.IsPrimaryAddress == "Y" || address.IsPrimaryAddress == null)
                                {
                                    customer_Address += address.Address ?? "";
                                    customer_Address += address.Ward != null ? ", " + address.Ward : "";
                                    customer_Address += address.District != null ? ", " + address.District : "";
                                    customer_Address += address.City != null ? ", " + address.City : "";
                                }
                            }
                        }
                        else if (hasORGANIZATION)
                        {
                            foreach (var address in item.ORGANIZATION_CUSTOMER.ADDRESS)
                            {
                                if (address.IsPrimaryAddress == "Y" || address.IsPrimaryAddress == null)
                                {
                                    customer_Address += address.Address ?? "";
                                    customer_Address += address.Ward != null ? ", " + address.Ward : "";
                                    customer_Address += address.District != null ? ", " + address.District : "";
                                    customer_Address += address.City != null ? ", " + address.City : "";
                                }
                            }
                        }

                        // Lấy số điện thoại khách hàng
                        if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.CONTACT.Mobile != null)
                        {
                            customer_Phone = item.INDIVIDUAL_CUSTOMER.CONTACT.Mobile;
                            isCustomer_Phone = "Y";
                        }
                        else if (hasORGANIZATION && item.ORGANIZATION_CUSTOMER.CONTACT.BusinessTelephone != null)
                        {
                            customer_Phone = item.ORGANIZATION_CUSTOMER.CONTACT.BusinessTelephone;
                            isCustomer_Phone = "Y";
                        }

                        // Lấy số fax khách hàng
                        if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.CONTACT.Fax != null)
                        {
                            customerFax = item.INDIVIDUAL_CUSTOMER.CONTACT.Fax;
                        }
                        else if (hasORGANIZATION && item.ORGANIZATION_CUSTOMER.CONTACT.Fax != null)
                        {
                            customerFax = item.ORGANIZATION_CUSTOMER.CONTACT.Fax;
                        }

                        // Lấy số CMND/CCC khách hàng
                        if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.IDNumber != null)
                        {
                            customerIDNumber = item.INDIVIDUAL_CUSTOMER.IDNumber;
                            isCustomer_IDNumber = "Y";
                        }

                        // Lấy mã số thuế khách hàng
                        if (hasORGANIZATION && item.ORGANIZATION_CUSTOMER.RegistrationNumber != null)
                        {
                            customerRegistrationNumber = item.ORGANIZATION_CUSTOMER.RegistrationNumber;
                            isCustomer_RegistrationNumber = "Y";
                        }

                        // Lấy tài khoản VAT khách hàng
                        if (hasORGANIZATION && item.ORGANIZATION_CUSTOMER.VATRegistrationNumber != null)
                        {
                            customerVATRegistrationNumber = item.ORGANIZATION_CUSTOMER.VATRegistrationNumber;
                            isCustomer_VATRegistrationNumber = "Y";
                        }

                        // Lấy tên người đại diện khách hàng
                        if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.CONTACT.ContactPerson != null)
                        {
                            customerRepresen = item.INDIVIDUAL_CUSTOMER.CONTACT.ContactPerson;

                            isCustomer_Represen = "N";
                        }
                        else if (hasREPRESENTATIVE && item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.RepresentativeName != null)
                        {
                            customerRepresen = item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.RepresentativeName;

                            isCustomer_Represen = "Y";
                        }

                        // Lấy chức vụ người đại diện khách hàng
                        if (hasREPRESENTATIVE && item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.Position != null)
                        {
                            customerPosition = item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.Position;

                            isCustomer_Position = "Y";
                        }
                        else if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.PartyName != null)
                        {
                            customerPosition = item.INDIVIDUAL_CUSTOMER.PartyName;

                            isCustomer_Position = "N";
                        }

                        // Lấy số giấy ủy quyền người đại diện khách hàng
                        if (hasREPRESENTATIVE && item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.AttorneyNo != null)
                        {
                            customer_AttorneyNo = item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.AttorneyNo;
                            isCustomer_AttorneyNo = "Y";
                        }

                        policy.Add("Customer_PartyNameVN", customer_PartyNameVN);
                        policy.Add("Customer_Address", customer_Address);
                        policy.Add("IsInsuredAndPolicyHolder", isInsuredAndPolicyHolder);
                        policy.Add("Customer_Phone", customer_Phone);
                        policy.Add("IsCustomer_Phone", isCustomer_Phone);
                        policy.Add("Customer_Fax", customerFax);
                        policy.Add("Customer_IDNumber", customerIDNumber);
                        policy.Add("IsCustomer_IDNumber", isCustomer_IDNumber);
                        policy.Add("Customer_RegistrationNumber", customerRegistrationNumber);
                        policy.Add("IsCustomer_RegistrationNumber", isCustomer_RegistrationNumber);
                        policy.Add("Customer_VATRegistrationNumber", customerVATRegistrationNumber);
                        policy.Add("IsCustomer_VATRegistrationNumber", isCustomer_VATRegistrationNumber);
                        policy.Add("Customer_Represen", customerRepresen);
                        policy.Add("IsCustomer_Represen", isCustomer_Represen);
                        policy.Add("Customer_Position", customerPosition);
                        policy.Add("IsCustomer_Position", isCustomer_Position);
                        policy.Add("Customer_AttorneyNo", customer_AttorneyNo);
                        policy.Add("IsCustomer_AttorneyNo", isCustomer_AttorneyNo);
                        policy.Add("IsIndividual", isIndividual);

                    }

                    // Lấy thông tin người được bảo hiểm
                    if (item.IsInsured == "Y")
                    {
                        string isCustomer_Phone_S = "N",
                            isCustomer_IDNumber_S = "N",
                            isCustomer_RegistrationNumber_S = "N",
                            isCustomer_VATRegistrationNumber_S = "N",
                            isCustomer_Represen_S = "N",
                            isCustomer_Position_S = "N",
                            isCustomer_AttorneyNo_S = "N";

                        // Kiểm tra người được bảo hiểm là cá nhân hay tổ chức
                        if (hasORGANIZATION)
                        {
                            isIndividual_S = "N";
                        }
                        else if (hasINDIVIDUAL)
                        {
                            isIndividual_S = "Y";
                        }

                        // Lấy tên người được bảo hiểm
                        if (hasINDIVIDUAL)
                        {
                            customer_PartyNameVN_S = item.INDIVIDUAL_CUSTOMER.PartyName;
                        }
                        else if (hasORGANIZATION)
                        {
                            customer_PartyNameVN_S = item.ORGANIZATION_CUSTOMER.PartyNameVN;
                        }

                        // Lấy địa chỉ người được bảo hiểm
                        if (hasINDIVIDUAL)
                        {
                            foreach (var address in item.INDIVIDUAL_CUSTOMER.ADDRESS)
                            {
                                if (address.IsPrimaryAddress == "Y" || address.IsPrimaryAddress == null)
                                {
                                    customer_Address_S += address.Address ?? "";
                                    customer_Address_S += address.Ward != null ? ", " + address.Ward : "";
                                    customer_Address_S += address.District != null ? ", " + address.District : "";
                                    customer_Address_S += address.City != null ? ", " + address.City : "";
                                }
                            }
                        }
                        else if (hasORGANIZATION)
                        {
                            foreach (var address in item.ORGANIZATION_CUSTOMER.ADDRESS)
                            {
                                if (address.IsPrimaryAddress == "Y" || address.IsPrimaryAddress == null)
                                {
                                    customer_Address_S += address.Address ?? "";
                                    customer_Address_S += address.Ward != null ? ", " + address.Ward : "";
                                    customer_Address_S += address.District != null ? ", " + address.District : "";
                                    customer_Address_S += address.City != null ? ", " + address.City : "";
                                }
                            }
                        }

                        // Lấy số điện thoại người được bảo hiểm
                        if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.CONTACT.Mobile != null)
                        {
                            customer_Phone_S = item.INDIVIDUAL_CUSTOMER.CONTACT.Mobile;
                            isCustomer_Phone_S = "Y";
                        }
                        else if (hasORGANIZATION && item.ORGANIZATION_CUSTOMER.CONTACT.BusinessTelephone != null)
                        {
                            customer_Phone_S = item.ORGANIZATION_CUSTOMER.CONTACT.BusinessTelephone;
                            isCustomer_Phone_S = "Y";
                        }

                        // Lấy số fax người được bảo hiểm
                        if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.CONTACT.Fax != null)
                        {
                            customerFax_S = item.INDIVIDUAL_CUSTOMER.CONTACT.Fax;
                        }
                        else if (hasORGANIZATION && item.ORGANIZATION_CUSTOMER.CONTACT.Fax != null)
                        {
                            customerFax_S = item.ORGANIZATION_CUSTOMER.CONTACT.Fax;
                        }

                        // Lấy số CMND/CCC người được bảo hiểm
                        if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.IDNumber != null)
                        {
                            customerIDNumber_S = item.INDIVIDUAL_CUSTOMER.IDNumber;
                            isCustomer_IDNumber_S = "Y";
                        }

                        // Lấy mã số thuế người được bảo hiểm
                        if (hasORGANIZATION && item.ORGANIZATION_CUSTOMER.RegistrationNumber != null)
                        {
                            customerRegistrationNumber_S = item.ORGANIZATION_CUSTOMER.RegistrationNumber;
                            isCustomer_RegistrationNumber_S = "Y";
                        }

                        // Lấy tài khoản VAT người được bảo hiểm
                        if (hasORGANIZATION && item.ORGANIZATION_CUSTOMER.VATRegistrationNumber != null)
                        {
                            customerVATRegistrationNumber_S = item.ORGANIZATION_CUSTOMER.VATRegistrationNumber;
                            isCustomer_VATRegistrationNumber_S = "Y";
                        }

                        // Lấy tên người đại diện người được bảo hiểm
                        if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.CONTACT.ContactPerson != null)
                        {
                            customerRepresen_S = item.INDIVIDUAL_CUSTOMER.CONTACT.ContactPerson;

                            isCustomer_Represen_S = "N";
                        }
                        else if (hasREPRESENTATIVE && item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.RepresentativeName != null)
                        {
                            customerRepresen_S = item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.RepresentativeName;

                            isCustomer_Represen_S = "Y";
                        }

                        // Lấy chức vụ người đại diện người được bảo hiểm
                        if (hasREPRESENTATIVE && item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.Position != null)
                        {
                            customerPosition_S = item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.Position;

                            isCustomer_Position_S = "Y";
                        }
                        else if (hasINDIVIDUAL && item.INDIVIDUAL_CUSTOMER.PartyName != null)
                        {
                            customerPosition_S = item.INDIVIDUAL_CUSTOMER.PartyName;

                            isCustomer_Position_S = "N";
                        }

                        // Lấy số giấy ủy quyền người đại diện người được bảo hiểm
                        if (hasREPRESENTATIVE && item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.AttorneyNo != null)
                        {
                            customer_AttorneyNo_S = item.ORGANIZATION_CUSTOMER.REPRESENTATIVE.AttorneyNo;
                            isCustomer_AttorneyNo_S = "Y";
                        }

                        policy.Add("Customer_PartyNameVN_S", customer_PartyNameVN_S);
                        policy.Add("Customer_Address_S", customer_Address_S);
                        policy.Add("Customer_Phone_S", customer_Phone_S);
                        policy.Add("Customer_Fax_S", customerFax_S);
                        policy.Add("IsIndividual_S", isIndividual_S);

                        policy.Add("IsCustomer_Phone_S", isCustomer_Phone_S);



                        policy.Add("Customer_IDNumber_S", customerIDNumber_S);
                        policy.Add("IsCustomer_IDNumber_S", isCustomer_IDNumber_S);
                        policy.Add("Customer_RegistrationNumber_S", customerRegistrationNumber_S);
                        policy.Add("IsCustomer_RegistrationNumber_S", isCustomer_RegistrationNumber_S);
                        policy.Add("Customer_VATRegistrationNumber_S", customerVATRegistrationNumber_S);
                        policy.Add("IsCustomer_VATRegistrationNumber_S", isCustomer_VATRegistrationNumber_S);
                        policy.Add("Customer_Represen_S", customerRepresen_S);
                        policy.Add("IsCustomer_Represen_S", isCustomer_Represen_S);
                        policy.Add("Customer_Position_S", customerPosition_S);
                        policy.Add("IsCustomer_Position_S", isCustomer_Position_S);
                        policy.Add("Customer_AttorneyNo_S", customer_AttorneyNo_S);
                        policy.Add("IsCustomer_AttorneyNo_S", isCustomer_AttorneyNo_S);
                    }
                }
            }

            // Hàm lấy thông tin rủi ro
            void GetRiskInfo(Dictionary<string, string> policy)
            {
                string hasNSUB_COVERAGE_BI = "N", // check > 1 hàng mục bảo hiểm BI
                        hasNSUB_COVERAGE = "N", // check > 1 hàng mục bảo hiểm
                        hasSUB_COVERAGE = "N"; // check = 1 hàng mục bảo hiểm

                hasCOVFCI = false;
                hasCOVFSP = false;
                hasCOVFPA = false;
                hasCOVFBI = false;
                string isCOVFBI = "N",
                       isCOVFSP = "N",
                       isCOVFPA = "N",
                       isCOVFSP_FPA = "N",
                       stt2 = "III",
                       stt = "IV",
                       sttNTH = "7",
                       sttPVBH = "";

                certificate.ForEach(item =>
                {
                    hasCOVFBI = item.RISKS.RISK.COVFBI != null ? true : false;
                    hasCOVFSP = item.RISKS.RISK.COVFSP != null ? true : false;
                    hasCOVFPA = item.RISKS.RISK.COVFPA != null ? true : false;
                    hasCOVFCI = item.RISKS.RISK.COVFCI != null ? true : false;
                });

                certificate.ForEach(item =>
                {
                    if (hasCOVFBI)
                    {
                        if (item.RISKS.RISK.COVFBI.CoverageCode != null)
                        {
                            isCOVFBI = "Y";
                            stt2 = "III";
                            stt = "IV";
                            sttNTH = "2";
                        }
                    }
                    if (hasCOVFSP)
                    {
                        if (item.RISKS.RISK.COVFSP.CoverageCode != null)
                        {
                            isCOVFSP = "Y";
                            sttPVBH = "a. ";
                            isCOVFSP_FPA = "Y";
                        }
                    }
                    if (hasCOVFPA)
                    {
                        if (item.RISKS.RISK.COVFPA.CoverageCode != null)
                        {
                            isCOVFPA = "Y";
                            sttPVBH = "a. ";
                            isCOVFSP_FPA = "Y";
                        }
                    }
                });

                certificate.ForEach(item =>
                {
                    if (hasCOVFSP)
                    {
                        if (item.RISKS.RISK.COVFSP.SUB_COVERAGE.Count > 1)
                        {
                            hasNSUB_COVERAGE = "Y";
                        }
                        if (item.RISKS.RISK.COVFSP.SUB_COVERAGE.Count > 0)
                        {
                            hasSUB_COVERAGE = "Y";
                        }
                    }
                    if (hasCOVFPA)
                    {
                        if (item.RISKS.RISK.COVFPA.SUB_COVERAGE.Count > 1)
                        {
                            hasNSUB_COVERAGE = "Y";
                        }
                        if (item.RISKS.RISK.COVFPA.SUB_COVERAGE.Count > 0)
                        {
                            hasSUB_COVERAGE = "Y";
                        }
                    }
                    if (hasCOVFBI)
                    {
                        if (item.RISKS.RISK.COVFBI.SUB_COVERAGE.Count > 1)
                        {
                            hasNSUB_COVERAGE_BI = "Y";
                        }
                    }
                });
                policy.Add("IsCOVFBI", isCOVFBI);
                policy.Add("IsCOVFSP", isCOVFSP);
                policy.Add("IsCOVFPA", isCOVFPA);
                policy.Add("IsCOVFSP_FPA", isCOVFSP_FPA);
                policy.Add("stt2", stt2);
                policy.Add("stt", stt);
                policy.Add("sttNTH", sttNTH);
                policy.Add("sttPVBH", sttPVBH);
                policy.Add("CountNRisk", certificate.Count > 1 ? "Y" : "N");
                policy.Add("CountNSUB_COVERAGE", hasNSUB_COVERAGE);
                policy.Add("CountNSUB_COVERAGE_BI", hasNSUB_COVERAGE_BI);
                policy.Add("HasSUB_COVERAGE", hasSUB_COVERAGE);

                // Kết thúc hàm lấy thông tin rủi ro
            }

            // Hàm lấy thông tin phí bảo hiểm
            void GetFeeInfo(Dictionary<string, string> policy)
            {
                string installmentDate = "............",
                             countNFEE_INSTALLMENT = "N";
                int countInstallmentPeriodSeq = 0;
                var policyFee = _root.GROUPPOLICY.POLICY_FEE_INFOR.SelectMany(i => i.POLICY_FEE).ToList();
                foreach (var item in policyFee)
                {
                    if (item.FeeType == "100101")
                    {
                        foreach (var i in item.POLICY_FEE_SPLIT)
                        {
                            foreach (var j in i.FEE_INSTALLMENT)
                            {
                                installmentDate = string.Format("{0:dd/MM/yyyy}", j.InstallmentDate);
                                countInstallmentPeriodSeq = (countInstallmentPeriodSeq < j.InstallmentPeriodSeq) ? j.InstallmentPeriodSeq : countInstallmentPeriodSeq; // max kỳ
                                countNFEE_INSTALLMENT = countInstallmentPeriodSeq > 1 ? "Y" : "N";
                            }
                        }
                    }
                }
                policy.Add("InstallmentDate", installmentDate);
                policy.Add("CountInstallmentPeriodSeq", countInstallmentPeriodSeq.ToString());
                policy.Add("CountNFEE_INSTALLMENT", countNFEE_INSTALLMENT);

                // Kết thúc hàm lấy thông tin phí bảo hiểm
            }

            // Hàm lấy thông tin phí bảo hiểm giấy chứng nhận
            void GetCertificateFeeInfo(Dictionary<string, string> policy)
            {
                decimal?
                          premiumGchFBI = 0, // Tổng phần 2
                          premiumGchFci = 0, // Cháy nổ bắt buộc
                          premiumGchFpaFsp = 0; // Rủi ro khác
                for (int j = 0; j < certificate.Count; j++)
                {
                    switch (productCode)
                    {
                        case "F01"
                            when hasCOVFCI:
                            premiumGchFci += certificate[j].RISKS.RISK.COVFCI.CoverageAnnualPremium;
                            premiumGchFBI += hasCOVFBI ? certificate[j].RISKS.RISK.COVFBI.CoverageAnnualPremium : 0;

                            // Mọi Rủi Ro Tài Sản
                            if (hasCOVFPA)
                            {
                                premiumGchFpaFsp += certificate[j].RISKS.RISK.COVFPA.CoverageAnnualPremium;
                            }
                            // Bảo hiểm hỏa hạn và các rủi ro đặc biệt
                            else if (hasCOVFSP)
                            {
                                premiumGchFpaFsp += certificate[j].RISKS.RISK.COVFSP.CoverageAnnualPremium;
                            }
                            break;
                        case "F02"
                            when hasCOVFSP:
                            premiumGchFBI += hasCOVFBI ? certificate[j].RISKS.RISK.COVFBI.CoverageAnnualPremium : 0;
                            // Mọi Rủi Ro Tài Sản
                            if (hasCOVFPA)
                            {
                                premiumGchFpaFsp += certificate[j].RISKS.RISK.COVFPA.CoverageAnnualPremium;
                            }
                            // Bảo hiểm hỏa hạn và các rủi ro đặc biệt
                            else if (hasCOVFSP)
                            {
                                premiumGchFpaFsp += certificate[j].RISKS.RISK.COVFSP.CoverageAnnualPremium;
                            }
                            break;
                    }
                }
                policy.Add("PremiumGchFciFpaFsp", (premiumGchFci + premiumGchFpaFsp).CurrencyS(currencyCode));
                policy.Add("PremiumGchFBI", premiumGchFBI.CurrencyS(currencyCode)); policy.Add("PremiumGchFci", premiumGchFci.CurrencyS(currencyCode));
                policy.Add("PremiumGchFpaFsp", premiumGchFpaFsp.CurrencyS(currencyCode));
            }

            // Hàm tạo đối tượng PolicyInformation từ dictionary policy
            PolicyInformation CreatePolicyInformation(Dictionary<string, string> policy)
            {
                PolicyInformation pI = new PolicyInformation();

                // Số
                pI.PolicyNo = (string.IsNullOrEmpty(groupPolicy.PolicyNo) && groupPolicy.IssuedDate == null) ? groupPolicy.QuotationNo : groupPolicy.PolicyNo;
                pI.QuotationNo = groupPolicy.QuotationNo;
                pI.IsClause = policy["IsClause"];
                // Hôm nay
                pI.CustomerConfirmationDate = string.Format("{0:dd} tháng {0:MM} năm {0:yyyy}", groupPolicy.CustomerConfirmationDate);
                // tại
                pI.BranchPlace = groupPolicy.GICBranchPlaceVN ?? "";
                // DOANH NGHIỆP BẢO HIỂM
                pI.GICBranchNameVN = groupPolicy.GICBranchNameVN?.ToUpper() ?? "";
                // Địa chỉ
                pI.GICBranchAddress = groupPolicy.GICBranchAddress ?? "";
                // Điện thoại
                pI.GICBranchPhone = groupPolicy.GICBranchPhone ?? "";
                // Fax
                pI.GICBranchFax = groupPolicy.GICBranchFax ?? "";
                // Mã số thuế
                pI.GICBranchTaxCode = groupPolicy.GICBranchTaxCode ?? "";
                pI.IsGICBranchBank = string.IsNullOrEmpty(groupPolicy.GICBranchBankAccount) ? "N" : "Y";
                // Tài khoản
                pI.GICBranchBank = string.IsNullOrEmpty(groupPolicy.GICBranchBankAccount) ? "" : $"{groupPolicy.GICBranchBankAccount} tại {groupPolicy.GICBranchBankName}";
                // Do Ông/Bà
                pI.GICRepresentative = groupPolicy.GICRepresentative ?? "";
                // Chức vụ
                pI.GICPositionRepresentative = groupPolicy.GICPositionRepresentative ?? "";
                pI.IsGICSignerContent = string.IsNullOrEmpty(groupPolicy.GICSignerContent) ? "N" : "Y";
                // Theo giấy uỷ quyền số 
                pI.GICSignerContent = groupPolicy.GICSignerContent ?? "";
                // NGƯỜI ĐƯỢC BẢO HIỂM
                pI.Customer_PartyNameVN = policy["Customer_PartyNameVN"];
                pI.Customer_PartyNameVN_S = policy["Customer_PartyNameVN_S"];
                // Địa chỉ
                pI.Customer_Address = policy["Customer_Address"];
                pI.Customer_Address_S = policy["Customer_Address_S"];
                pI.IsCustomer_Phone = policy["IsCustomer_Phone"];
                pI.IsCustomer_Phone_S = policy["IsCustomer_Phone_S"];
                // Điện thoại
                pI.Customer_Phone = policy["Customer_Phone"];
                pI.Customer_Phone_S = policy["Customer_Phone_S"];
                // Fax
                pI.Customer_Fax = policy["Customer_Fax"];
                pI.Customer_Fax_S = policy["Customer_Fax_S"];
                pI.IsCustomer_IDNumber = policy["IsCustomer_IDNumber"];
                pI.IsCustomer_IDNumber_S = policy["IsCustomer_IDNumber_S"];
                // CMND/CCC
                pI.Customer_IDNumber = policy["Customer_IDNumber"];
                pI.Customer_IDNumber_S = policy["Customer_IDNumber_S"];
                pI.IsCustomer_RegistrationNumber = policy["IsCustomer_RegistrationNumber"];
                pI.IsCustomer_RegistrationNumber_S = policy["IsCustomer_RegistrationNumber_S"];
                // Mã số thuế
                pI.Customer_RegistrationNumber = policy["Customer_RegistrationNumber"];
                pI.Customer_RegistrationNumber_S = policy["Customer_RegistrationNumber_S"];
                pI.IsCustomer_VATRegistrationNumber = policy["IsCustomer_VATRegistrationNumber"];
                pI.IsCustomer_VATRegistrationNumber_S = policy["IsCustomer_VATRegistrationNumber_S"];
                // Tài khoản
                pI.Customer_VATRegistrationNumber = policy["Customer_VATRegistrationNumber"];
                pI.Customer_VATRegistrationNumber_S = policy["Customer_VATRegistrationNumber_S"];
                pI.IsCustomer_Represen = policy["IsCustomer_Represen"];
                pI.IsCustomer_Represen_S = policy["IsCustomer_Represen_S"];
                // Do Ông/Bà
                pI.Customer_Represen = policy["Customer_Represen"];
                pI.Customer_Represen_S = policy["Customer_Represen_S"];
                pI.IsCustomer_Position = policy["IsCustomer_Position"];
                pI.IsCustomer_Position_S = policy["IsCustomer_Position_S"];
                // Chức vụ
                pI.Customer_Position = policy["Customer_Position"];
                pI.Customer_Position_S = policy["Customer_Position_S"];
                pI.IsCustomer_AttorneyNo = policy["IsCustomer_AttorneyNo"];
                pI.IsCustomer_AttorneyNo_S = policy["IsCustomer_AttorneyNo_S"];
                // Theo giấy uỷ quyền số
                pI.Customer_AttorneyNo = policy["Customer_AttorneyNo"];
                pI.Customer_AttorneyNo_S = policy["Customer_AttorneyNo_S"];
                pI.IsBeneficiary = groupPolicy.Beneficiary != null ? "Y" : "N";
                // NGƯỜI THỤ HƯỞNG
                pI.Beneficiary_Name = policy["Beneficiary_Name"];
                pI.Beneficiary_Detail = policy["Beneficiary_Detail"];
                pI.Beneficiary_Detail2 = policy["Beneficiary_Detail2"];
                // GĐKD
                pI.IsCOVFBI = policy["IsCOVFBI"];
                pI.CountNRisk = policy["CountNRisk"];
                pI.IsClauseDKBS = policy["IsClauseDKBS"];
                pI.CountNSUB_COVERAGE = policy["CountNSUB_COVERAGE"];
                pI.HasSUB_COVERAGE = policy["HasSUB_COVERAGE"];
                pI.IsClauseNTH = policy["IsClauseNTH"];
                pI.IsBi = policy["IsBi"];
                pI.CountNSUB_COVERAGE_BI = policy["CountNSUB_COVERAGE_BI"];
                // TỔNG PHÍ BẢO HIỂM THANH TOÁN
                pI.TotalPremium = groupPolicy.TotalPremium != null ? groupPolicy.TotalPremium.CurrencyS(currencyCode) : "";
                pI.TotalPremiumAfterVAT = groupPolicy.TotalPremiumAfterVAT != null ? groupPolicy.TotalPremiumAfterVAT.CurrencyS(currencyCode) : "";
                pI.CountNFEE_INSTALLMENT = policy["CountNFEE_INSTALLMENT"];
                // TỔNG VAT
                pI.VATAmount = groupPolicy.VATAmount != null ? groupPolicy.VATAmount.CurrencyS(currencyCode) : "";
                // ngày << nên cho vào vòng lặp policy_fee
                pI.InstallmentDate = policy["InstallmentDate"];
                // kỳ
                pI.CountInstallmentPeriodSeq = policy["CountInstallmentPeriodSeq"];
                pI.IsGICBranchBankAccount = policy["IsGICBranchBankAccount"];
                // Số tài khoản
                pI.GICBranchBankAccount = policy["GICBranchBankAccount"];
                pI.IsGICBranchBankName = groupPolicy.GICBranchBankName != null ? "Y" : "N";
                // Ngân hàng
                pI.GICBranchBankName = groupPolicy.GICBranchBankName ?? "";
                // GĐKD
                pI.stt = policy["stt"];
                pI.stt2 = policy["stt2"];
                pI.sttNTH = policy["sttNTH"];
                pI.stt4 = policy["stt4"];
                pI.stt5 = policy["stt5"];
                pI.stt6 = policy["stt6"];
                pI.stt7 = policy["stt7"];
                pI.stt8 = policy["stt8"];
                pI.stt9 = policy["stt9"];
                pI.IsCOVFPA = policy["IsCOVFPA"];
                pI.IsCOVFSP = policy["IsCOVFSP"];
                pI.IsCOVFSP_FPA = policy["IsCOVFSP_FPA"];
                pI.sttPVBH = policy["sttPVBH"];
                pI.CurrencyCode = currencyCode;
                pI.IsInsuredAndPolicyHolder = policy["IsInsuredAndPolicyHolder"];
                // GCN
                pI.PremiumGchFciFpaFsp = policy["PremiumGchFciFpaFsp"];
                pI.PremiumGchFBI = policy["PremiumGchFBI"];
                pI.PremiumGchFci = policy["PremiumGchFci"];
                pI.PremiumGchFpaFsp = policy["PremiumGchFpaFsp"];
                pI.IsIndividual = policy["IsIndividual"];
                pI.IsIndividual_S = policy["IsIndividual_S"];

                return pI;// Trả về đối tượng PolicyInformation
            }

            


            #region ~~~CCHD
            List<Clause> ClauseContent()
            {
                //var lst_Clause = new List<Clause>();

                //foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                //{
                //    if ((string.IsNullOrEmpty(item.PrintOrder) == false && item.PrintOrder == "CCHD") ||
                //        (string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && item.ClauseCoverageCodes == "CCHD"))
                //    {
                //        lst_Clause.Add(new Clause
                //        {
                //            ClauseContent = item.ClauseContent.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                                                                                                                                       ")
                //            ,
                //            ClauseCoverageCodes = item.ClauseCoverageCodes != null ? item.ClauseCoverageCodes : ""
                //        });
                //    }
                //}
                //lst_Clause = lst_Clause
                //  .OrderBy(x => x.ClauseCoverageCodes == "0" ? int.MaxValue : int.TryParse(x.ClauseCoverageCodes.TrimEnd("bic".ToCharArray()), out int result) ? result : int.MaxValue)
                //  .ToList();
                //for (int i = 0; i < lst_Clause.Count; i++)
                //{
                //    lst_Clause[i].STT = (i + 1);
                //}
                //return lst_Clause;

                var lst_Clause = new List<Clause>();

                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if ((!string.IsNullOrWhiteSpace(item.PrintOrder) && item.PrintOrder == "CCHD") ||
                        (!string.IsNullOrWhiteSpace(item.ClauseCoverageCodes) && item.ClauseCoverageCodes == "CCHD"))
                    {
                        lst_Clause.Add(new Clause
                        {
                            ClauseContent = item.ClauseContent.Replace("<br/>", Environment.NewLine + " "),
                            ClauseCoverageCodes = item.ClauseCoverageCodes ?? "0"
                            ,
                            PrintOrder = item.PrintOrder
                        });
                    }
                }
                lst_Clause = lst_Clause
                    //.OrderBy(p => p.PrintOrder == "0" ? int.MaxValue : int.Parse(p.PrintOrder.Remove(1, p.PrintOrder.Length - 1)))
                    .OrderBy(p => (p.PrintOrder ?? "0") == "0" ? int.MaxValue : int.Parse((p.PrintOrder ?? "0").Remove(1, (p.PrintOrder ?? "0").Length - 1)))
                    .ToList();
                for (int i = 0; i < lst_Clause.Count; i++)
                {
                    lst_Clause[i].STT = (i + 1);
                }
                return lst_Clause;
            }
            _dataModelView.addtable(DataDictionaryExtension.ConvertToDictionary(ClauseContent()), "~~~CCHD", 0);
            #endregion

            #region ~~~DKBS
            List<Clause> ClauseDKBS()
            {
                var lst_Clause = new List<Clause>();
                var stt = 1;
                var clauseLimit = "";

                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if ((string.IsNullOrEmpty(item.PrintOrder) == false && (item.PrintOrder == "CCHD" || item.PrintOrder == "BI" || item.PrintOrder == "NTH")) ||
                        (string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && (item.ClauseCoverageCodes.ToLower().Contains("cchd") || item.ClauseCoverageCodes.ToLower().Contains("bi") || item.ClauseCoverageCodes.ToLower().Contains("nth")))
                        )
                    {
                    }
                    else
                    {
                        if (item.ClauseCoverageCodes == null)
                        {
                            item.ClauseCoverageCodes = "0";
                        }
                        if (item.ClauseLimit != null && item.ClauseExecss != null)
                        {
                            if (groupPolicy.CurrencyCode == "vnd")
                            {
                                clauseLimit = "(" + item.ClauseLimit.Replace(',', '.') + "," + item.ClauseExecss + ")";
                            }
                            else
                            {
                                clauseLimit = "(" + item.ClauseLimit + "," + item.ClauseExecss + ")";
                            }
                        }
                        else if (item.ClauseLimit != null)
                        {
                            if (groupPolicy.CurrencyCode == "vnd")
                            {
                                clauseLimit = "(" + item.ClauseLimit.Replace(",0", ".0") + ")";
                            }
                            else
                            {
                                clauseLimit = "(" + item.ClauseLimit + ")";
                            }
                        }
                        else
                        {
                            clauseLimit = "";
                        }

                        lst_Clause.Add(new Clause
                        {
                            //ClauseContent = "3." + stt + " " + item.ClauseTitle + " " + clauseLimit
                            ClauseContent = "" + item.ClauseTitle + " " + clauseLimit
                            ,
                            ClauseCoverageCodes = item.ClauseCoverageCodes ?? "0"
                            ,
                            ClauseTitle = item.ClauseTitle + " " + clauseLimit
                            ,
                            PrintOrder = item.PrintOrder
                            ,
                            ClauseContent2 = item.ClauseContent.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                                                                                                                                       ")
                            ,
                            STT = stt
                        });
                        ++stt;
                    }
                }
                stt = 1;

                // Sử dụng OrderBy để trả về một danh sách mới đã được sắp xếp
                var sortedList = lst_Clause.OrderBy(x => x.PrintOrder == null ? int.MaxValue : Int32.Parse(x.PrintOrder)).ToList();

                for (int i = 0; i < sortedList.Count; i++)
                {
                    sortedList[i].STT = (i + 1);
                }
                return sortedList;
            }
            _dataModelView.addtable(DataDictionaryExtension.ConvertToDictionary(ClauseDKBS()), "~~~DKBS", 0);

            //// do trên file word cần canh đều, nên xuống dòng khi có khoản trắng dài gây ra lỗi. 
            /// ~~~DKBS được bổ sung bằng ~~ClauseAttachMaster và ~~ClauseAttachChild của Nhung, để có thể canh đều bình thường
            List<Clause> ClauseAttachMaster()
            {
                var lst_Clause = new List<Clause>();
                var stt = 1;
                var clauseLimit = "";

                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if ((string.IsNullOrEmpty(item.PrintOrder) == false && (item.PrintOrder == "CCHD" || item.PrintOrder == "BI" || item.PrintOrder == "NTH")) ||
                        (string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && (item.ClauseCoverageCodes.ToLower().Contains("cchd") || item.ClauseCoverageCodes.ToLower().Contains("bi") || item.ClauseCoverageCodes.ToLower().Contains("nth")))
                        )
                    {
                    }
                    else
                    {
                        if (item.ClauseCoverageCodes == null)
                        {
                            item.ClauseCoverageCodes = "0";
                        }
                        if (item.ClauseLimit != null && item.ClauseExecss != null)
                        {
                            if (groupPolicy.CurrencyCode == "vnd")
                            {
                                clauseLimit = "(" + item.ClauseLimit.Replace(',', '.') + "," + item.ClauseExecss + ")";
                            }
                            else
                            {
                                clauseLimit = "(" + item.ClauseLimit + "," + item.ClauseExecss + ")";
                            }
                        }
                        else if (item.ClauseLimit != null)
                        {
                            if (groupPolicy.CurrencyCode == "vnd")
                            {
                                clauseLimit = "(" + item.ClauseLimit.Replace(",0", ".0") + ")";
                            }
                            else
                            {
                                clauseLimit = "(" + item.ClauseLimit + ")";
                            }
                        }
                        else
                        {
                            clauseLimit = "";
                        }

                        lst_Clause.Add(new Clause
                        {
                            //ClauseContent = "3." + stt + " " + item.ClauseTitle + " " + clauseLimit
                            ClauseContent = "" + item.ClauseTitle + " " + clauseLimit
                            ,
                            ClauseCoverageCodes = item.ClauseCoverageCodes ?? "0"
                            ,
                            ClauseTitle = item.ClauseTitle + " " + clauseLimit
                            ,
                            ClauseContent2 = item.ClauseContent.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                                                                                                                                       ")
                            ,
                            ClauseContent3 = item.ClauseContent
                            ,
                            PrintOrder = item.PrintOrder
                            ,
                            KeyChildTable = (stt).ToString()
                            ,
                            STT = stt
                        });
                        ++stt;
                    }
                }
                stt = 1;

                // Sử dụng OrderBy để trả về một danh sách mới đã được sắp xếp
                var sortedList = lst_Clause.OrderBy(x => x.PrintOrder == null ? int.MaxValue : Int32.Parse(x.PrintOrder)).ToList();

                for (int i = 0; i < sortedList.Count; i++)
                {
                    sortedList[i].STT = (i + 1);
                }
                return sortedList;
            }
            List<Clause> ClauseAttachChild()
            {
                var lst_Clause = new List<Clause>();
                var parentList = ClauseAttachMaster();
                foreach (var item in parentList)
                {
                    var benefitStr = item.ClauseContent3.Split(new string[] { "<br/>" }, StringSplitOptions.None);
                    for (int j = 0; j < benefitStr.Length; j++)
                    {
                        if (!string.IsNullOrEmpty(benefitStr[j]))
                        {
                            lst_Clause.Add(new Clause
                            {
                                ClauseTitle = item.ClauseTitle,
                                ClauseContent = benefitStr[j],
                                PrintOrder = item.PrintOrder,
                                KeyChildTable = item.KeyChildTable,
                            });
                        }
                    }
                }
                var sortedList = lst_Clause.OrderBy(x => x.PrintOrder == null ? int.MaxValue : Int32.Parse(x.PrintOrder)).ToList();

                for (int i = 0; i < sortedList.Count; i++)
                {
                    sortedList[i].STT = (i + 1);
                }
                return sortedList;
            }

            _dataModelView.addtable_lv2(DataDictionaryExtension.ConvertToDictionary(ClauseAttachMaster()), DataDictionaryExtension.ConvertToDictionary(ClauseAttachChild()), "KeyChildTable", "~~~ClauseAttachMaster", "~~~ClauseAttachChild", 0);
            #endregion

            #region ~~~DKBSBI
            List<Clause> ClauseDKBSBI()
            {
                var lst_Clause = new List<Clause>();
                var stt = 1;
                var clauseLimit = "";

                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if ((string.IsNullOrWhiteSpace(item.PrintOrder) == false && (item.PrintOrder == "BI" || item.PrintOrder == "C")) ||
                        (string.IsNullOrWhiteSpace(item.ClauseCoverageCodes) == false && (item.ClauseCoverageCodes.ToLower().Contains("bi") || item.ClauseCoverageCodes.ToLower().Contains("c"))))
                    {
                        if (item.ClauseCoverageCodes == null)
                        {
                            item.ClauseCoverageCodes = "0";
                        }
                        if (item.ClauseLimit != null && item.ClauseExecss != null)
                        {
                            if (groupPolicy.CurrencyCode == "vnd")
                            {
                                clauseLimit = "(" + item.ClauseLimit.Replace(',', '.') + "," + item.ClauseExecss + ")";
                            }
                            else
                            {
                                clauseLimit = "(" + item.ClauseLimit + "," + item.ClauseExecss + ")";
                            }
                        }
                        else if (item.ClauseLimit != null)
                        {
                            if (groupPolicy.CurrencyCode == "vnd")
                            {
                                clauseLimit = "(" + item.ClauseLimit.Replace(",0", ".0") + ")";
                            }
                            else
                            {
                                clauseLimit = "(" + item.ClauseLimit + ")";
                            }
                        }
                        else
                        {
                            clauseLimit = "";
                        }

                        lst_Clause.Add(new Clause
                        {
                            //ClauseContent = "3." + stt + " " + item.ClauseTitle + " " + clauseLimit
                            ClauseContent = "" + item.ClauseTitle + " " + clauseLimit
                            ,
                            ClauseCoverageCodes = item.ClauseCoverageCodes ?? "0"
                            //phung.huynhminh add
                            ,
                            STT = stt
                            ,
                            PrintOrder = item.PrintOrder
                            ,
                            ClauseContent2 = item.ClauseTitle + " " + clauseLimit
                        });
                        ++stt;
                    }
                    // Sort the list in order
                }
                stt = 1;
                lst_Clause = lst_Clause
                    //.OrderBy(p => p.PrintOrder == "0" ? int.MaxValue : int.Parse(p.PrintOrder.Remove(1, p.PrintOrder.Length - 1)))
                    .OrderBy(p => (p.PrintOrder ?? "0") == "0" ? int.MaxValue : int.Parse((p.PrintOrder ?? "0").Remove(1, (p.PrintOrder ?? "0").Length - 1)))
                    .ToList();
                for (int i = 0; i < lst_Clause.Count; i++)
                {
                    lst_Clause[i].STT = (i + 1);
                }
                return lst_Clause;
            }
            _dataModelView.addtable(DataDictionaryExtension.ConvertToDictionary(ClauseDKBSBI()), "~~~DKBSBI", 0);
            #endregion

            #region ~~~NTH
            List<Clause> ClauseNTH()
            {
                var lst_ClauseNTH = new List<Clause>();

                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if (_root.GROUPPOLICY.QuotationDate != null)
                    {
                        if ((string.IsNullOrEmpty(item.PrintOrder) == false && item.PrintOrder == "NTH") ||
                            (string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && item.ClauseCoverageCodes == "NTH"))
                        {
                            lst_ClauseNTH.Add(new Clause
                            {
                                ClauseContent = item.ClauseContent.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                                                                                                     ")
                                ,
                                PrintOrder = item.PrintOrder
                                ,
                                ClauseCoverageCodes = item.ClauseCoverageCodes ?? "0"
                            });
                        }
                    }

                }
                lst_ClauseNTH = lst_ClauseNTH
                   //.OrderBy(p => p.PrintOrder == "0" ? int.MaxValue : int.Parse(p.PrintOrder.Remove(1, p.PrintOrder.Length - 1)))
                    .OrderBy(p => (p.PrintOrder ?? "0") == "0" ? int.MaxValue : int.Parse((p.PrintOrder ?? "0").Remove(1, (p.PrintOrder ?? "0").Length - 1)))
                   .ToList();
                for (int i = 0; i < lst_ClauseNTH.Count; i++)
                {
                    lst_ClauseNTH[i].STT = (i + 1);
                }
                return lst_ClauseNTH;
            }
            _dataModelView.addtable(DataDictionaryExtension.ConvertToDictionary(ClauseNTH()), "~~~NTH", 0);

            //// do trên file word cần canh đều, nên xuống dòng khi có khoản trắng dài gây ra lỗi. 
            /// ~~~NTH được bổ sung bằng ~~NTHAttachMaster và ~~NTHAttachChild của Nhung, để có thể canh đều bình thường
            List<Clause> NTHAttachMaster()
            {
                var lst_Clause = new List<Clause>();
                foreach (var item in groupPolicy.CLAUSES.SelectMany(cs => cs.CLAUSE))
                {
                    if (_root.GROUPPOLICY.QuotationDate != null)
                    {
                        if ((string.IsNullOrEmpty(item.PrintOrder) == false && item.PrintOrder == "NTH") ||
                            (string.IsNullOrEmpty(item.ClauseCoverageCodes) == false && item.ClauseCoverageCodes == "NTH"))
                        {
                            lst_Clause.Add(new Clause
                            {
                                ClauseContent = item.ClauseContent.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                                                                                                     ")
                                ,
                                PrintOrder = item.PrintOrder
                                ,
                                ClauseContent3 = item.ClauseContent
                                ,
                                ClauseCoverageCodes = item.ClauseCoverageCodes ?? "0"
                            });
                        }
                    }

                }
                // Sử dụng OrderBy để trả về một danh sách mới đã được sắp xếp
                var sortedList = lst_Clause.OrderBy(x => x.PrintOrder == null ? int.MaxValue : Int32.Parse(x.PrintOrder)).ToList();
                for (int i = 0; i < sortedList.Count; i++)
                {
                    sortedList[i].STT = (i + 1);
                }
                return sortedList;
            }
            List<Clause> NTHAttachChild()
            {
                var lst_Clause = new List<Clause>();
                var parentList = NTHAttachMaster();
                foreach (var item in parentList)
                {
                    var benefitStr = item.ClauseContent3.Split(new string[] { "<br/>" }, StringSplitOptions.None);
                    for (int j = 0; j < benefitStr.Length; j++)
                    {
                        if (!string.IsNullOrEmpty(benefitStr[j]))
                        {
                            lst_Clause.Add(new Clause
                            {
                                ClauseTitle = item.ClauseTitle,
                                ClauseContent = benefitStr[j],
                                PrintOrder = item.PrintOrder,
                                KeyChildTable = item.KeyChildTable,
                            });
                        }
                    }
                }
                var sortedList = lst_Clause.OrderBy(x => x.PrintOrder == null ? int.MaxValue : Int32.Parse(x.PrintOrder)).ToList();

                for (int i = 0; i < sortedList.Count; i++)
                {
                    sortedList[i].STT = (i + 1);
                }
                return sortedList;
            }

            _dataModelView.addtable_lv2(DataDictionaryExtension.ConvertToDictionary(NTHAttachMaster()), DataDictionaryExtension.ConvertToDictionary(NTHAttachChild()), "KeyChildTable", "~~~NTHAttachMaster", "~~~NTHAttachChild", 0);
            #endregion

            #region ~~~InsuredSubCoverageName
            List<tb_InsuredSubCoverageName> InsuredSubCoverageName()
            {
                var lst_InsuredSubCoverageName = new List<tb_InsuredSubCoverageName>();
                for (int i = 0; i < 1; i++)
                {
                    if (hasCOVFSP)
                    {
                        if (productCode == "F02")
                        {
                            foreach (var item in certificate[0].RISKS.RISK.COVFSP.FIRE_SPECAL_PERILS)
                            {
                                var subCoverageLimit = "";
                                if (item.SubLimit != null)
                                {
                                    //subCoverageLimit = "(" + item.SubCoverageLimit.CurrencyS(currencyCode) + ")";
                                    subCoverageLimit = "(" + item.SubLimit + ")";
                                }
                                var _ = item.SubCoverageName == "A. Hoả hoạn, sét đánh, Nổ nồi hơi,  hơi đốt phục vụ sinh hoạt trong nhà" ? " A. Cháy, Sét đánh và Nổ của nồi hơi dân dụng" : item.SubCoverageName;
                                lst_InsuredSubCoverageName.Add(new tb_InsuredSubCoverageName
                                {

                                    SubCoverageName = _,
                                    SubCoverageLimit = subCoverageLimit
                                });
                            }
                        }
                        if (productCode == "F01")
                        {
                            foreach (var item in certificate[0].RISKS.RISK.COVFSP.POLICY_PERIL)
                            {
                                var subCoverageLimit = "";
                                if (item.SubCoverageLimit != null)
                                {
                                    //subCoverageLimit = "(" + item.SubCoverageLimit.CurrencyS(currencyCode) + ")";
                                    subCoverageLimit = "(" + item.SubCoverageLimit + ")";
                                }
                                var _ = item.SubCoverageName == "A. Hoả hoạn, sét đánh, Nổ nồi hơi,  hơi đốt phục vụ sinh hoạt trong nhà" ? " A. Cháy, Sét đánh và Nổ của nồi hơi dân dụng" : item.SubCoverageName;
                                lst_InsuredSubCoverageName.Add(new tb_InsuredSubCoverageName
                                {
                                    SubCoverageName = _,
                                    SubCoverageLimit = subCoverageLimit
                                });
                            }
                        }
                    }
                }
                return lst_InsuredSubCoverageName.OrderBy(o => o.SubCoverageName).ToList();
            }
            _dataModelView.addtable(DataDictionaryExtension.ConvertToDictionary(InsuredSubCoverageName()), "~~~InsuredSubCoverageName", 0);
            #endregion

            #region ~~~POLICY_FEE
            List<tb_POLICY_FEE> PolicyFee()
            {
                var policyFee = _root.GROUPPOLICY.POLICY_FEE_INFOR.SelectMany(i => i.POLICY_FEE).ToList();
                var lstPolicyFee = new List<tb_POLICY_FEE>();

                foreach (var item in policyFee)
                {
                    if (item.FeeType == "100101")
                    {
                        foreach (var i in item.POLICY_FEE_SPLIT)
                        {
                            foreach (var j in i.FEE_INSTALLMENT)
                            {
                                lstPolicyFee.Add(new tb_POLICY_FEE
                                {
                                    InstallmentDate = string.Format("{0:dd/MM/yyyy}", j.InstallmentDate),
                                    InstallmentPeriodSeq = j.InstallmentPeriodSeq,
                                    InstallmentAmount = j.InstallmentAmount.CurrencyS(currencyCode),
                                    CurrencyCode = currencyCode,
                                });
                            }
                        }
                    }
                }
                return lstPolicyFee;
            }
            _dataModelView.addtable(DataDictionaryExtension.ConvertToDictionary(PolicyFee()), "~~~POLICY_FEE", 0);
            #endregion

            switch (productCode)
            {
                case "F01":
                    #region ~~~InsuredF1
                    List<tb_Insured> InsuredContentF1()
                    {
                        var lst_Insured = new List<tb_Insured>();
                        for (int i = 0; i < certificate.Count; i++)
                        {
                            var EffectiveDate = certificate[i].EffectiveDate;
                            var ExpiryDate = certificate[i].ExpiryDate;
                            var typeOfRiskCompulsoryCodes = certificate[i].RISKS.RISK.TypeOfRiskCompulsoryCode == null ? "" : certificate[i].RISKS.RISK.TypeOfRiskCompulsoryCode;
                            // The converted string array
                            var typeOfRiskCompulsoryCode = typeOfRiskCompulsoryCodes.Split(new[] { '.' }, 2);
                            lst_Insured.Add(new tb_Insured
                            {
                                STT = (i + 1).ToString(),
                                STTRoman = NumberToString.IntToRoman(i + 1),
                                InsuredObject = certificate[i].RISKS.RISK.InsuredObject.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                    "),
                                LocationinsuredPrintedinform = certificate[i].RISKS.RISK.LocationinsuredPrintedinform, //  Địa điểm bảo hiểm
                                Occupancy = certificate[i].RISKS.RISK.Occupancy, //  Mục đích sử dụng
                                //TypeOfRiskCompulsoryCode = certificate[i].RISKS.RISK.TypeOfRiskCompulsoryCode.Substring(0, 1), //  Thuộc danh mục cơ sở ex: 1 đến 9
                                TypeOfRiskCompulsoryCode = typeOfRiskCompulsoryCode[0], //  Thuộc danh mục cơ sở
                            });
                        }
                        return lst_Insured;
                    }
                    _dataModelView.addtable(DataDictionaryExtension.ConvertToDictionary(InsuredContentF1()), "~~~InsuredF1", 0);
                    #endregion

                    #region ~~~InsuredNF1FCI ~~~InsuredSubNF1FCI
                    List<tb_Insured> InsuredContentNF1FCI()
                    {
                        var lst_Insured = new List<tb_Insured>();
                        for (int i = 0; i < certificate.Count; i++)
                        {
                            var EffectiveDate = certificate[i].EffectiveDate;
                            var ExpiryDate = certificate[i].ExpiryDate;
                            // The original string
                            var typeOfRiskCompulsoryCodes = certificate[i].RISKS.RISK.TypeOfRiskCompulsoryCode == null ? "" : certificate[i].RISKS.RISK.TypeOfRiskCompulsoryCode;
                            // The converted string array
                            var typeOfRiskCompulsoryCode = typeOfRiskCompulsoryCodes.Split(new[] { '.' }, 2);
                            if (hasCOVFCI)
                            {
                                //var _CoverageTotalSumInsuredToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.COVFCI.CoverageTotalSumInsuredToWord.Split('|')[1] : certificate[i].RISKS.RISK.COVFCI.CoverageTotalSumInsuredToWord.Split('|')[0];
                                var _CoverageTotalSumInsured = certificate[i].RISKS.RISK.COVFCI.CoverageTotalSumInsured;
                                var _CoverageTotalSumInsuredToWord = classReadMoney.VietNamConvert(_CoverageTotalSumInsured.ToString(), currencyCode);
                                lst_Insured.Add(new tb_Insured
                                {
                                    STT = (i + 1).ToString(),
                                    STTRoman = NumberToString.IntToRoman(i + 1),

                                    InsuredObject = certificate[i].RISKS.RISK.InsuredObject.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                    "),
                                    LocationinsuredPrintedinform = certificate[i].RISKS.RISK.LocationinsuredPrintedinform, //  Địa điểm bảo hiểm
                                    Occupancy = certificate[i].RISKS.RISK.Occupancy, //  Mục đích sử dụng
                                    TypeOfRiskCompulsoryCode = typeOfRiskCompulsoryCode[0], //  Thuộc danh mục cơ sở

                                    SubCoverageSumInsured = certificate[i].RISKS.RISK.COVFCI.SUB_COVERAGE[0].SubCoverageSumInsured.CurrencyS(currencyCode),
                                    CurrencyCode = currencyCode,
                                    CoverageTotalSumInsuredToWord = _CoverageTotalSumInsuredToWord,
                                    DeductibleDescription = certificate[i].RISKS.RISK.COVFCI.DEDUCTIBLE != null ? ": " + certificate[i].RISKS.RISK.COVFCI.DEDUCTIBLE.DeductibleDescription.Replace("|", Environment.NewLine + "                                                                                                                                                       ") : "",
                                    DeductibleDescription2 = certificate[i].RISKS.RISK.COVFCI.DEDUCTIBLE != null ? certificate[i].RISKS.RISK.COVFCI.DEDUCTIBLE.DeductibleDescription.Replace("|", Environment.NewLine + "                                                                                                                                                       ") : "",
                                    EffectiveExpiryDate = string.Format("Từ {0:HH} giờ {0:mm} ngày {0:dd/MM/yyyy} đến {1:HH} giờ {1:mm} ngày {1:dd/MM/yyyy}", EffectiveDate, ExpiryDate),
                                    CoverageTotalSumInsured = _CoverageTotalSumInsured.CurrencyS(currencyCode),
                                    // thời hạn bồi thường tối đa
                                    IndemnityPeriod = certificate[i].RISKS.RISK.COVFCI.IndemnityPeriod + " " + (certificate[i].RISKS.RISK.COVFCI.IndemnityPeriodMode == "M" ? "tháng" : "năm"),

                                    KeyChildTable = (i + 1).ToString()
                                });
                            }
                        }
                        return lst_Insured;
                    }
                    List<tb_Insured> InsuredContentSubNF1FCI()
                    {
                        var lst_Insured = new List<tb_Insured>();
                        for (int j = 0; j < certificate.Count; j++)
                        {
                            var _stt = 1;

                            if (hasCOVFCI)
                            {
                                string[] separatingStrings = { "<br/>" };

                                var _CoveragePremiumBeforeAdjusted = certificate[j].RISKS.RISK.COVFCI.CoveragePremiumBeforeAdjusted.CurrencyS(currencyCode);
                                var _CoverageVATAmount = certificate[j].RISKS.RISK.COVFCI.CoverageVATAmount.CurrencyS(currencyCode);
                                var _CoveragePremiumAfterTax = certificate[j].RISKS.RISK.COVFCI.CoveragePremiumAfterTax.CurrencyS(currencyCode);

                                foreach (var i in certificate[j].RISKS.RISK.COVFCI.SUB_COVERAGE)
                                {
                                    lst_Insured.Add(new tb_Insured
                                    {
                                        STT = _stt.ToString(),
                                        SubCoverageName = i.SubCoverageName, //  Hạng mục được bảo hiểm
                                        SubCoverageSumInsured = i.SubCoverageSumInsured.CurrencyS(currencyCode), //  Số tiền bảo hiểm 

                                        KeyChildTable = (j + 1).ToString(),
                                    });
                                    _stt++;
                                }
                            }
                        }
                        return lst_Insured;
                    }

                    _dataModelView.addtable_lv2(DataDictionaryExtension.ConvertToDictionary(InsuredContentNF1FCI()), DataDictionaryExtension.ConvertToDictionary(InsuredContentSubNF1FCI()), "KeyChildTable", "~~~InsuredNF1FCI", "~~~InsuredSubNF1FCI", 0);
                    #endregion

                    #region ~~~InsuredNF1FBI ~~~InsuredSubNF1FBI
                    List<tb_Insured> InsuredNF1FBI()
                    {
                        var lst_Insured = new List<tb_Insured>();
                        for (int i = 0; i < certificate.Count; i++)
                        {
                            if (hasCOVFBI)
                            {
                                //var _TotalPremiumWord = currencyCode == "VND" ? _root.GROUPPOLICY.TotalPremiumAfterVATInword.Split('|')[1] : _root.GROUPPOLICY.TotalPremiumAfterVATInword.Split('|')[0];
                                var _TotalPremiumWord = classReadMoney.VietNamConvert(_root.GROUPPOLICY.TotalPremium.ToString(), currencyCode);
                                var _CoveragePremiumAfterTax = certificate[i].RISKS.RISK.COVFBI.CoveragePremiumAfterTax;
                                //var _CoveragePremiumAfterTaxToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.COVFBI.CoveragePremiumAfterTaxToWord.Split('|')[1] : certificate[i].RISKS.RISK.COVFBI.CoveragePremiumAfterTaxToWord.Split('|')[0];
                                var _CoveragePremiumAfterTaxToWord = classReadMoney.VietNamConvert(certificate[i].RISKS.RISK.COVFBI.CoveragePremiumAfterTax.ToString(), currencyCode);
                                var _CoverageTotalSumInsured = certificate[i].RISKS.RISK.COVFBI.CoverageTotalSumInsured;
                                //var _IndemnityPeriodMode = certificate[i].RISKS.RISK.COVFBI.IndemnityPeriodMode == "Y" ? "năm" : "tháng";
                                //_IndemnityPeriodMode = certificate[i].RISKS.RISK.COVFBI.IndemnityPeriodMode == "M" ? "tháng" : "ngày";
                                var _EffectiveDate = _root.CERTIFICATES.CERTIFICATE[i].EffectiveDate;
                                var _ExpiryDate = _root.CERTIFICATES.CERTIFICATE[i].ExpiryDate;

                                //var _CoverageTotalSumInsuredToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.COVFBI.CoverageTotalSumInsuredToWord.Split('|')[1] : certificate[i].RISKS.RISK.COVFBI.CoverageTotalSumInsuredToWord.Split('|')[0];
                                var _CoverageTotalSumInsuredToWord = classReadMoney.VietNamConvert(certificate[i].RISKS.RISK.COVFBI.SUB_COVERAGE[0].SubCoverageSumInsured.ToString(), currencyCode);

                                lst_Insured.Add(new tb_Insured
                                {
                                    STT = (i + 1).ToString(),
                                    STTRoman = NumberToString.IntToRoman(i + 1),
                                    CurrencyCode = currencyCode,
                                    CoverageTotalSumInsured = _CoverageTotalSumInsured.CurrencyS(currencyCode),
                                    CoverageTotalSumInsuredToWord = _CoverageTotalSumInsuredToWord,
                                    //IndemnityPeriod = certificate[i].RISKS.RISK.COVFBI.IndemnityPeriod,
                                    //IndemnityPeriodMode = _IndemnityPeriodMode,
                                    DeductibleDescription = certificate[i].RISKS.RISK.COVFBI.DEDUCTIBLE != null ? certificate[i].RISKS.RISK.COVFBI.DEDUCTIBLE.DeductibleDescription.Replace("|", Environment.NewLine + "                                                                                                                                                                                                                                                                                       ") : "",
                                    EffectiveExpiryDate = string.Format("Từ {0:HH} giờ {0:mm} ngày {0:dd/MM/yyyy} đến {1:HH} giờ {1:mm} ngày {1:dd/MM/yyyy}", _EffectiveDate, _ExpiryDate),
                                    LocationinsuredPrintedinform = certificate[i].RISKS.RISK.LocationinsuredPrintedinform, //  Địa điểm bảo hiểm
                                                                                                                           // thời hạn bồi thường tối đa
                                    IndemnityPeriod = certificate[i].RISKS.RISK.COVFBI.IndemnityPeriod + " " + (certificate[i].RISKS.RISK.COVFBI.IndemnityPeriodMode == "M" ? "tháng" : "năm"),
                                    KeyChildTable = (i + 1).ToString()
                                });
                            }
                        }
                        return lst_Insured;
                    }
                    List<tb_Insured> InsuredSubNF1FBI()
                    {
                        var lst_Insured = new List<tb_Insured>();

                        for (int j = 0; j < certificate.Count; j++)
                        {
                            var _stt = 1;
                            if (hasCOVFBI)
                            {
                                foreach (var i in certificate[j].RISKS.RISK.COVFBI.SUB_COVERAGE)
                                {
                                    lst_Insured.Add(new tb_Insured
                                    {
                                        STT = _stt.ToString(),
                                        // Hạng mục được bảo hiểm
                                        SubCoverageCode = i.SubCoverageCode,
                                        SubCoverageName = i.SubCoverageName,
                                        // Số tiền bảo hiểm
                                        SubCoverageSumInsured = i.SubCoverageSumInsured.CurrencyS(currencyCode),
                                        KeyChildTable = (j + 1).ToString(),
                                    });
                                    _stt++;
                                }
                            }
                        }
                        lst_Insured = lst_Insured.OrderBy(i => i.KeyChildTable).ThenBy(i => i.SubCoverageCode).ToList();
                        var _KeyChildTable = "";
                        int _j = 1;
                        for (int i = 0; i < lst_Insured.Count; i++)
                        {
                            // lần đầu tiên thì gián
                            if (i == 0)
                                _KeyChildTable = lst_Insured[i].KeyChildTable;
                            // reset to 1
                            if (_KeyChildTable != lst_Insured[i].KeyChildTable)
                            {
                                _KeyChildTable = lst_Insured[i].KeyChildTable;
                                _j = 1;
                            }
                            // 1
                            lst_Insured[i].STT = (_j).ToString();
                            // 2
                            _j++;

                        }
                        return lst_Insured;


                        //var lstInsuredSort = new List<tb_Insured>();
                        //var num = 1;
                        //foreach (var cl in lst_Insured)
                        //{
                        //    lstInsuredSort.Add(new tb_Insured
                        //    {
                        //        STT = num.ToString(),
                        //        SubCoverageCode = cl.SubCoverageCode,
                        //        // Hạng mục được bảo hiểm
                        //        SubCoverageName = cl.SubCoverageName,
                        //        SubCoverageSumInsured = cl.SubCoverageSumInsured,
                        //        // Số tiền bảo hiểm 
                        //        KeyChildTable = cl.KeyChildTable,
                        //      });
                        //    num++;
                        //}
                        //return lstInsuredSort;
                    }
                    _dataModelView.addtable_lv2(DataDictionaryExtension.ConvertToDictionary(InsuredNF1FBI()), DataDictionaryExtension.ConvertToDictionary(InsuredSubNF1FBI()), "KeyChildTable", "~~~InsuredNF1FBI", "~~~InsuredSubNF1FBI", 0);
                    #endregion

                    #region ~~~InsuredNF1S ~~~InsuredSubNF1S
                    List<tb_Insured> InsuredNF1S()
                    {
                        var lstInsured = new List<tb_Insured>();
                        for (int i = 0; i < certificate.Count; i++)
                        {
                            lstInsured.Add(new tb_Insured
                            {
                                STT = (i + 1).ToString(),
                                STTRoman = NumberToString.IntToRoman(i + 1),
                                //  Tổng 
                                CoveragePremiumAfterTax = certificate[i].RISKS.RISK.RiskTotalGrossPremium.CurrencyS(currencyCode),
                                CurrencyCode = currencyCode,
                                //  Tổng cộng 
                                //CoveragePremiumAfterTaxToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.RiskTotalGrossPremium_LCToWord.Split('|')[1] : certificate[i].RISKS.RISK.RiskTotalGrossPremium_LCToWord.Split('|')[0],
                                CoveragePremiumAfterTaxToWord = classReadMoney.VietNamConvert(certificate[i].RISKS.RISK.RiskTotalGrossPremium_LC.ToString(), currencyCode),
                                KeyChildTable = (i + 1).ToString()
                            });
                        }
                        return lstInsured;
                    }
                    List<tb_Insured> InsuredSubNF1S()
                    {
                        var lstInsured = new List<tb_Insured>();
                        var otherPremiumAfterTax = (decimal?)0;

                        var newLine = "                                                                                                                                                       ";

                        for (int j = 0; j < certificate.Count; j++)
                        {
                            var stt = 1;
                            decimal? subCoverageAnnualRateFci = 0, // Tỷ lệ phí Cháy nổ bắt buộc 
                                        subCoveragePremiumFci = 0, // Phí Cháy nổ bắt buộc 
                                        subCoverageVATFci = 0, // VAT Cháy nổ bắt buộc 
                                        subCoverageAnnualRateOther = 0, // Tỷ lệ phí Các rủi ro khác
                                        subCoveragePremiumOther = 0, // Phí Các rủi ro khác
                                        subCoverageVATOther = 0, // VAT Các rủi ro khác
                                        subCoverageAnnualRateS = 0, // Tỷ lệ phí Gián đoạn kinh doanh
                                        subCoveragePremiumS = 0, // Phí Gián đoạn kinh doanh
                                        subCoverageVATAmountS = 0, // VAT Gián đoạn kinh doanh
                                        subCoveragePremiumAfterTaxS = 0; // Tổng phí phần II
                            var infoInsured = "";
                            if (hasCOVFCI)
                            {
                                var covfciCount = certificate[j].RISKS.RISK.COVFCI.SUB_COVERAGE.Count();
                                /*
                                    Sum các giá trị Cháy Nổ Bắt Buộc:
                                    subCoverageAnnualRateFci Tỷ lệ phí
                                    subCoveragePremiumFci Phí
                                    subCoveragePremiumFci VAT
                                */
                                for (int l = 0; l < covfciCount; l++)
                                {
                                    var fci = certificate[j].RISKS.RISK.COVFCI.SUB_COVERAGE[l];
                                    subCoverageAnnualRateFci = fci.SubCoverageCompulsoryAnnualRate;
                                    subCoveragePremiumFci += fci.SubCoveragePremium;
                                    subCoverageVATFci = (subCoveragePremiumFci * 10 / 100);
                                }
                                // Có gián đoạn kinh doanh
                                if (hasCOVFBI)
                                {
                                    var covFbiCount = certificate[j].RISKS.RISK.COVFBI.SUB_COVERAGE.Count();
                                    /*     
                                       sum các giá trị Gián đoạn kinh doanh:
                                       subCoverageAnnualRateS Tỷ lệ phí Gián đoạn kinh doanh
                                       subCoveragePremiumS Phí Gián đoạn kinh doanh
                                       subCoverageVATAmountS VAT Gián đoạn kinh doanh
                                       subCoveragePremiumAfterTaxS Tổng phí phần II
                                    */
                                    for (int l = 0; l < covFbiCount; l++)
                                    {
                                        var fbi = certificate[j].RISKS.RISK.COVFBI.SUB_COVERAGE[l];
                                        subCoverageAnnualRateS = fbi.SubCoverageAnnualRate;
                                        subCoveragePremiumS += fbi.SubCoveragePremium;
                                        subCoverageVATAmountS = (subCoveragePremiumS * 10 / 100);
                                        subCoveragePremiumAfterTaxS = (subCoveragePremiumS + subCoverageVATAmountS);
                                    }
                                }
                                // nếu không có rủi ro khác, ngoài cháy nổ bắt buộc
                                if (hasCOVFSP != true && hasCOVFPA != true)
                                {
                                    infoInsured += "Tỷ lệ phí bảo hiểm     : " + String.Format("{0:0.##}", subCoverageAnnualRateFci * 100).Replace(".", ",") + " % (chưa bao gồm VAT)" + "                                                                                            ";
                                    infoInsured += "Phí bảo hiểm              : " + subCoveragePremiumFci.CurrencyS(currencyCode) + " " + currencyCode + newLine;
                                    infoInsured += "VAT                           : " + subCoverageVATFci.CurrencyS(currencyCode) + " " + currencyCode + newLine;
                                    lstInsured.Add(new tb_Insured
                                    {
                                        SubCoveragePremiumAfterTax = (subCoveragePremiumFci + subCoverageVATFci).CurrencyS(currencyCode),
                                        CurrencyCode = currencyCode,
                                        InfoInsured = infoInsured,
                                        KeyChildTable = (j + 1).ToString()
                                    });
                                }
                                // có mở rộng thêm Hỏa hoạn và các rủi ro đặc biệt hoặc Mọi rủi ro tài sản
                                else
                                {
                                    var f = new SUB_COVERAGE();
                                    // Mọi Rủi Ro Tài Sản
                                    if (hasCOVFPA)
                                    {
                                        var _covfpa_count = certificate[j].RISKS.RISK.COVFPA.SUB_COVERAGE.Count();
                                        // sum các giá trị covfpa
                                        for (int l = 0; l < _covfpa_count; l++)
                                        {
                                            f = certificate[j].RISKS.RISK.COVFPA.SUB_COVERAGE[l];
                                            subCoverageAnnualRateOther = f.SubCoverageVoluntaryAdditionalRate;
                                            subCoveragePremiumOther += f.SubCoverageVoluntaryAdditionalPremium;
                                            subCoverageVATOther = (subCoveragePremiumOther * 10 / 100);
                                        }
                                    }
                                    // Bảo hiểm hỏa hạn và các rủi ro đặc biệt
                                    else if (hasCOVFSP)
                                    {
                                        var covFspCount = certificate[j].RISKS.RISK.COVFSP.SUB_COVERAGE.Count();
                                        // sum các giá trị covfsp
                                        for (int l = 0; l < covFspCount; l++)
                                        {
                                            f = certificate[j].RISKS.RISK.COVFSP.SUB_COVERAGE[l];
                                            subCoverageAnnualRateOther = f.SubCoverageCompulsoryAnnualRate;
                                            subCoveragePremiumOther += f.SubCoveragePremium;
                                            subCoverageVATOther = (subCoveragePremiumOther * 10 / 100);
                                        }
                                    }
                                    // Tỷ lệ phí bảo hiểm 
                                    infoInsured += "Tỷ lệ phí bảo hiểm          : " + String.Format("{0:0.##}", (subCoverageAnnualRateFci + subCoverageAnnualRateOther) * 100).Replace(".", ",") + " % (chưa bao gồm VAT)" + newLine;
                                    infoInsured += "a. Cháy nổ bắt buộc        : " + String.Format("{0:0.##}", subCoverageAnnualRateFci * 100).Replace(".", ",") + " % (chưa bao gồm VAT)" + newLine;
                                    infoInsured += "b. Các rủi ro khác           : " + String.Format("{0:0.##}", subCoverageAnnualRateOther * 100).Replace(".", ",") + " % (chưa bao gồm VAT)" + newLine;
                                    // Phí bảo hiểm
                                    infoInsured += "Phí bảo hiểm:" + newLine;
                                    infoInsured += "a. Cháy nổ bắt buộc        : " + subCoveragePremiumFci.CurrencyS(currencyCode) + " " + currencyCode + newLine;
                                    infoInsured += "b. Các rủi ro khác           : " + subCoveragePremiumOther.CurrencyS(currencyCode) + " " + currencyCode + newLine;
                                    infoInsured += "Phí bảo hiểm (a+b)         : " + (subCoveragePremiumFci + subCoveragePremiumOther).CurrencyS(currencyCode) + " " + currencyCode + newLine;
                                    // VAT
                                    infoInsured += "VAT                               : " + (subCoverageVATFci + subCoverageVATOther).CurrencyS(currencyCode) + " " + currencyCode + newLine;
                                    // Tổng phí
                                    otherPremiumAfterTax = (subCoveragePremiumFci + subCoveragePremiumOther) + (subCoverageVATFci + subCoverageVATOther);
                                    lstInsured.Add(new tb_Insured
                                    {
                                        InfoInsured = infoInsured,
                                        STTRoman = NumberToString.IntToRoman(stt),
                                        SubCoveragePremiumAfterTax = otherPremiumAfterTax.CurrencyS(currencyCode), // Tổng phí phần I = (_f.SubCoveragePremium * 10 / 100).CurrencyS(_CurrencyCode), // VAT
                                        //// --Gián đoạn kinh doanh--
                                        // Tỷ lệ phí bảo hiểm
                                        SubCoverageAnnualRateS = (subCoverageAnnualRateS * 100).ToString(),
                                        // Phí bảo hiểm
                                        SubCoveragePremiumS = subCoveragePremiumS.CurrencyS(currencyCode),
                                        // VAT
                                        SubCoverageVATAmountS = (subCoverageVATAmountS).CurrencyS(currencyCode),
                                        // Tổng phí phần II
                                        SubCoveragePremiumAfterTaxS = subCoveragePremiumAfterTaxS.CurrencyS(currencyCode),
                                        CurrencyCode = currencyCode,
                                        KeyChildTable = (j + 1).ToString()
                                    });
                                    stt++;
                                }
                            }
                        }
                        return lstInsured;
                    }
                    _dataModelView.addtable_lv2(DataDictionaryExtension.ConvertToDictionary(InsuredNF1S()), DataDictionaryExtension.ConvertToDictionary(InsuredSubNF1S()), "KeyChildTable", "~~~InsuredNF1S", "~~~InsuredSubNF1S", 0);
                    #endregion

                    break;
                case "F02":
                    #region ~~~Insured
                    List<tb_Insured> Insured()
                    {
                        var lstInsured = new List<tb_Insured>();
                        for (int i = 0; i < certificate.Count; i++)
                        {
                            var effectiveDate = certificate[i].EffectiveDate;
                            var expiryDate = certificate[i].ExpiryDate;
                            var typeOfRiskCompulsoryCodes = certificate[i].RISKS.RISK.TypeOfRiskCompulsoryCode == null ? "" : certificate[i].RISKS.RISK.TypeOfRiskCompulsoryCode;
                            // The converted string array
                            var typeOfRiskCompulsoryCode = typeOfRiskCompulsoryCodes.Split(new[] { '.' }, 2);
                            if (hasCOVFSP)
                            {
                                //var coverageTotalSumInsuredToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.COVFSP.CoverageTotalSumInsuredToWord.Split('|')[1] : certificate[i].RISKS.RISK.COVFSP.CoverageTotalSumInsuredToWord.Split('|')[0];
                                var coverageTotalSumInsuredToWord = classReadMoney.VietNamConvert(certificate[i].RISKS.RISK.COVFSP.CoverageTotalSumInsured.ToString(), currencyCode);

                                //var totalPremiumWord = currencyCode == "VND" ? _root.GROUPPOLICY.TotalPremiumAfterVATInword.Split('|')[1] : _root.GROUPPOLICY.TotalPremiumAfterVATInword.Split('|')[0];
                                var totalPremiumWord = classReadMoney.VietNamConvert(_root.GROUPPOLICY.TotalPremiumAfterVAT.ToString(), currencyCode);

                                var coveragePremiumAfterTax = certificate[i].RISKS.RISK.COVFSP.CoveragePremiumAfterTax;
                                //var coveragePremiumAfterTaxToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.COVFSP.CoveragePremiumAfterTaxToWord.Split('|')[1] : certificate[i].RISKS.RISK.COVFSP.CoveragePremiumAfterTaxToWord.Split('|')[0];
                                var coveragePremiumAfterTaxToWord = classReadMoney.VietNamConvert(certificate[i].RISKS.RISK.COVFSP.CoveragePremiumAfterTax.ToString(), currencyCode);

                                var coverageTotalSumInsured = certificate[i].RISKS.RISK.COVFSP.CoverageTotalSumInsured;

                                lstInsured.Add(new tb_Insured
                                {
                                    STT = (i + 1).ToString(),
                                    STTRoman = NumberToString.IntToRoman(i + 1),
                                    LocationinsuredPrintedinform = certificate[i].RISKS.RISK.LocationinsuredPrintedinform,
                                    InsuredObject = certificate[i].RISKS.RISK.InsuredObject.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                    "),
                                    Occupancy = certificate[i].RISKS.RISK.Occupancy,
                                    CurrencyCode = currencyCode,
                                    SubCoverageSumInsured = certificate[i].RISKS.RISK.COVFSP.SUB_COVERAGE[0].SubCoverageSumInsured.CurrencyS(currencyCode),
                                    CoverageTotalSumInsured = coverageTotalSumInsured.CurrencyS(currencyCode),
                                    CoverageTotalSumInsuredToWord = coverageTotalSumInsuredToWord,
                                    DeductibleDescription = certificate[i].RISKS.RISK.COVFSP.DEDUCTIBLE != null ? certificate[i].RISKS.RISK.COVFSP.DEDUCTIBLE.DeductibleDescription.Replace("|", Environment.NewLine + "                                                                                                                                                                                                                                                                                       ") : "",
                                    EffectiveExpiryDate = string.Format("Từ {0:HH} giờ {0:mm} ngày {0:dd/MM/yyyy} đến {1:HH} giờ {1:mm} ngày {1:dd/MM/yyyy}", effectiveDate, expiryDate),
                                    // Tổng 
                                    CoveragePremiumAfterTax = coveragePremiumAfterTax.CurrencyS(currencyCode),
                                    // Tổng cộng 
                                    CoveragePremiumAfterTaxToWord = coveragePremiumAfterTaxToWord,
                                    //  Thuộc danh mục cơ sở
                                    TypeOfRiskCompulsoryCode = typeOfRiskCompulsoryCode[0],
                                    KeyChildTable = (i + 1).ToString()
                                });
                            }
                        }
                        return lstInsured;
                    }
                    _dataModelView.addtable(DataDictionaryExtension.ConvertToDictionary(Insured()), "~~~Insured", 0);
                    #endregion

                    #region ~~~InsuredN ~~~InsuredSubN
                    List<tb_Insured> InsuredSubN()
                    {
                        var lstInsured = new List<tb_Insured>();
                        for (int j = 0; j < certificate.Count; j++)
                        {
                            var effectiveDate = certificate[j].EffectiveDate;
                            var expiryDate = certificate[j].ExpiryDate;
                            if (hasCOVFSP)
                            {
                                var stt = 1;
                                var coveragePremiumBeforeAdjusted = certificate[j].RISKS.RISK.COVFSP.CoveragePremiumBeforeAdjusted.CurrencyS(currencyCode);
                                var coverageVATAmount = certificate[j].RISKS.RISK.COVFSP.CoverageVATAmount.CurrencyS(currencyCode);
                                var coveragePremiumAfterTax = certificate[j].RISKS.RISK.COVFSP.CoveragePremiumAfterTax.CurrencyS(currencyCode);
                                var coverageTotalSumInsuredToWord = classReadMoney.VietNamConvert(certificate[j].RISKS.RISK.COVFSP.CoverageTotalSumInsured_LC.ToString(), currencyCode);

                                foreach (var i in certificate[j].RISKS.RISK.COVFSP.SUB_COVERAGE)
                                {
                                    lstInsured.Add(new tb_Insured
                                    {
                                        STT = stt.ToString(),
                                        STTRoman = NumberToString.IntToRoman(stt),
                                        CoverageTotalSumInsuredToWord = coverageTotalSumInsuredToWord,
                                        DeductibleDescription = certificate[j].RISKS.RISK.COVFSP.DEDUCTIBLE != null ? certificate[j].RISKS.RISK.COVFSP.DEDUCTIBLE.DeductibleDescription.Replace("|", Environment.NewLine + "                                                                                                                                                                                                                                                                                       ") : "",
                                        EffectiveExpiryDate = string.Format("Từ {0:HH} giờ {0:mm} ngày {0:dd/MM/yyyy} đến {1:HH} giờ {1:mm} ngày {1:dd/MM/yyyy}", effectiveDate, expiryDate),
                                        // Hạng mục được bảo hiểm
                                        SubCoverageName = i.SubCoverageName,
                                        // Số tiền bảo hiểm 
                                        SubCoverageSumInsured = i.SubCoverageSumInsured.CurrencyS(currencyCode),
                                        // Phí bảo hiểm
                                        SubCoveragePremium = i.SubCoveragePremium.CurrencyS(currencyCode),
                                        // Tỷ lệ phí bảo hiểm
                                        SubCoverageAnnualRate = (i.SubCoverageAnnualRate * 100).ToString(),
                                        // VAT
                                        SubCoverageVATAmount = (i.SubCoveragePremium * 10 / 100).CurrencyS(currencyCode),
                                        // Tổng phí 
                                        SubCoveragePremiumAfterTax = (i.SubCoveragePremium + (i.SubCoveragePremium * 10 / 100)).CurrencyS(currencyCode),
                                        KeyChildTable = (j + 1).ToString()
                                    });
                                    stt++;
                                }
                            }
                        }
                        return lstInsured;
                    }
                    _dataModelView.addtable_lv2(DataDictionaryExtension.ConvertToDictionary(Insured()), DataDictionaryExtension.ConvertToDictionary(InsuredSubN()), "KeyChildTable", "~~~InsuredN", "~~~InsuredSubN", 0);
                    #endregion

                    #region ~~~InsuredNFBI ~~~InsuredSubNFBI
                    List<tb_Insured> InsuredFBI()
                    {
                        var lstInsured = new List<tb_Insured>();
                        for (int i = 0; i < certificate.Count; i++)
                        {
                            var effectiveDate = certificate[i].EffectiveDate;
                            var expiryDate = certificate[i].ExpiryDate;
                            if (hasCOVFBI)
                            {
                                //var coverageTotalSumInsuredToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.COVFBI.CoverageTotalSumInsuredToWord.Split('|')[1] : certificate[i].RISKS.RISK.COVFBI.CoverageTotalSumInsuredToWord.Split('|')[0];
                                var coverageTotalSumInsuredToWord = classReadMoney.VietNamConvert(certificate[i].RISKS.RISK.COVFBI.CoverageTotalSumInsured.ToString(), currencyCode);

                                //var totalPremiumWord = currencyCode == "VND" ? _root.GROUPPOLICY.TotalPremiumAfterVATInword.Split('|')[1] : _root.GROUPPOLICY.TotalPremiumAfterVATInword.Split('|')[0];
                                var totalPremiumWord = classReadMoney.VietNamConvert(_root.GROUPPOLICY.TotalPremiumAfterVAT.ToString(), currencyCode);

                                var coveragePremiumAfterTax = certificate[i].RISKS.RISK.COVFBI.CoveragePremiumAfterTax;
                                //var coveragePremiumAfterTaxToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.COVFBI.CoveragePremiumAfterTaxToWord.Split('|')[1] : certificate[i].RISKS.RISK.COVFBI.CoveragePremiumAfterTaxToWord.Split('|')[0];
                                var coveragePremiumAfterTaxToWord = classReadMoney.VietNamConvert(certificate[i].RISKS.RISK.COVFBI.CoveragePremiumAfterTax.ToString(), currencyCode);

                                var coverageTotalSumInsured = certificate[i].RISKS.RISK.COVFBI.CoverageTotalSumInsured;

                                lstInsured.Add(new tb_Insured
                                {
                                    STT = (i + 1).ToString(),
                                    STTRoman = NumberToString.IntToRoman(i + 1),
                                    LocationinsuredPrintedinform = certificate[i].RISKS.RISK.LocationinsuredPrintedinform,
                                    InsuredObject = certificate[i].RISKS.RISK.InsuredObject.Replace("<br/>", Environment.NewLine + "                                                                                                                                                                    "),
                                    Occupancy = certificate[i].RISKS.RISK.Occupancy,
                                    CurrencyCode = currencyCode,
                                    SubCoverageSumInsured = certificate[i].RISKS.RISK.COVFBI.SUB_COVERAGE[0].SubCoverageSumInsured.CurrencyS(currencyCode),
                                    CoverageTotalSumInsured = coverageTotalSumInsured.CurrencyS(currencyCode),
                                    CoverageTotalSumInsuredToWord = coverageTotalSumInsuredToWord,
                                    DeductibleDescription = certificate[i].RISKS.RISK.COVFBI.DEDUCTIBLE != null ? certificate[i].RISKS.RISK.COVFBI.DEDUCTIBLE.DeductibleDescription.Replace("|", Environment.NewLine + "                                                                                                                                                                                                                                                                                       ") : "",
                                    EffectiveExpiryDate = string.Format("Từ {0:HH} giờ {0:mm} ngày {0:dd/MM/yyyy} đến {1:HH} giờ {1:mm} ngày {1:dd/MM/yyyy}", effectiveDate, expiryDate),
                                    // Tổng 
                                    CoveragePremiumAfterTax = coveragePremiumAfterTax.CurrencyS(currencyCode),
                                    // Tổng cộng 
                                    CoveragePremiumAfterTaxToWord = coveragePremiumAfterTaxToWord,

                                    KeyChildTable = (i + 1).ToString()
                                });
                            }
                        }
                        return lstInsured;
                    }
                    List<tb_Insured> InsuredSubNFBI()
                    {
                        var lstInsured = new List<tb_Insured>();
                        for (int j = 0; j < certificate.Count; j++)
                        {
                            var effectiveDate = certificate[j].EffectiveDate;
                            var expiryDate = certificate[j].ExpiryDate;
                            if (hasCOVFBI)
                            {
                                var stt = 1;
                                var coveragePremiumBeforeAdjusted = certificate[j].RISKS.RISK.COVFBI.CoveragePremiumBeforeAdjusted.CurrencyS(currencyCode);
                                var coverageVATAmount = certificate[j].RISKS.RISK.COVFBI.CoverageVATAmount.CurrencyS(currencyCode);
                                var coveragePremiumAfterTax = certificate[j].RISKS.RISK.COVFBI.CoveragePremiumAfterTax.CurrencyS(currencyCode);

                                foreach (var i in certificate[j].RISKS.RISK.COVFBI.SUB_COVERAGE)
                                {
                                    lstInsured.Add(new tb_Insured
                                    {
                                        STT = stt.ToString(),
                                        STTRoman = NumberToString.IntToRoman(stt),
                                        // Hạng mục được bảo hiểm
                                        SubCoverageCode = i.SubCoverageCode,
                                        SubCoverageName = i.SubCoverageName,
                                        // Số tiền bảo hiểm 
                                        SubCoverageSumInsured = i.SubCoverageSumInsured.CurrencyS(currencyCode),
                                        // Phí bảo hiểm
                                        SubCoveragePremium = i.SubCoveragePremium.CurrencyS(currencyCode),
                                        // Tỷ lệ phí bảo hiểm
                                        SubCoverageAnnualRate = (i.SubCoverageAnnualRate * 100).ToString(),
                                        // VAT
                                        SubCoverageVATAmount = (i.SubCoveragePremium * 10 / 100).CurrencyS(currencyCode),
                                        // Tổng phí 
                                        SubCoveragePremiumAfterTax = (i.SubCoveragePremium + (i.SubCoveragePremium * 10 / 100)).CurrencyS(currencyCode),
                                        KeyChildTable = (j + 1).ToString()
                                    });
                                    stt++;
                                }
                            }
                        }
                        lstInsured = lstInsured.OrderBy(i => i.KeyChildTable).ThenBy(i => i.SubCoverageCode).ToList();
                        var _KeyChildTable = "";
                        int _j = 1;
                        // Printer: 1,2,3,4 : 1,2,3,4
                        for (int i = 0; i < lstInsured.Count; i++)
                        {
                            // lần đầu tiên thì gián
                            if (i == 0)
                                _KeyChildTable = lstInsured[i].KeyChildTable;
                            // reset to 1
                            if (_KeyChildTable != lstInsured[i].KeyChildTable)
                            {
                                _KeyChildTable = lstInsured[i].KeyChildTable;
                                _j = 1;
                            }
                            // 1
                            lstInsured[i].STT = (_j).ToString();
                            // 2
                            _j++;

                        }
                        return lstInsured;
                    }
                    _dataModelView.addtable_lv2(DataDictionaryExtension.ConvertToDictionary(InsuredFBI()), DataDictionaryExtension.ConvertToDictionary(InsuredSubNFBI()), "KeyChildTable", "~~~InsuredNFBI", "~~~InsuredSubNFBI", 0);
                    #endregion

                    #region ~~~InsuredNF2S ~~~InsuredSubNF2S
                    List<tb_Insured> InsuredNF2S()
                    {
                        var lstInsured = new List<tb_Insured>();
                        for (int i = 0; i < certificate.Count; i++)
                        {
                            lstInsured.Add(new tb_Insured
                            {
                                STT = (i + 1).ToString(),
                                STTRoman = NumberToString.IntToRoman(i + 1),
                                //  Tổng 
                                CoveragePremiumAfterTax = certificate[i].RISKS.RISK.RiskTotalGrossPremium.CurrencyS(currencyCode),
                                CurrencyCode = currencyCode,
                                //  Tổng cộng 
                                //CoveragePremiumAfterTaxToWord = currencyCode == "VND" ? certificate[i].RISKS.RISK.RiskTotalGrossPremium_LCToWord.Split('|')[1] : certificate[i].RISKS.RISK.RiskTotalGrossPremium_LCToWord.Split('|')[0],
                                CoveragePremiumAfterTaxToWord = classReadMoney.VietNamConvert(certificate[i].RISKS.RISK.RiskTotalGrossPremium_LC.ToString(), currencyCode),

                                KeyChildTable = (i + 1).ToString()
                            });
                        }
                        return lstInsured;
                    }
                    List<tb_Insured> InsuredSubNF2S()
                    {
                        var lstInsured = new List<tb_Insured>();
                        var otherPremiumAfterTax = (decimal?)0;

                        var newLine = "                                                                                                                                                       ";

                        for (int j = 0; j < certificate.Count; j++)
                        {
                            var stt = 1;
                            decimal? subCoverageAnnualRateFci = 0, // Tỷ lệ phí Cháy nổ bắt buộc 
                                        subCoveragePremiumFci = 0, // Phí Cháy nổ bắt buộc 
                                        subCoverageVATFci = 0, // VAT Cháy nổ bắt buộc 
                                        subCoverageAnnualRateOther = 0, // Tỷ lệ phí Các rủi ro khác
                                        subCoveragePremiumOther = 0, // Phí Các rủi ro khác
                                        subCoverageVATOther = 0, // VAT Các rủi ro khác
                                        subCoverageAnnualRateS = 0, // Tỷ lệ phí Gián đoạn kinh doanh
                                        subCoveragePremiumS = 0, // Phí Gián đoạn kinh doanh
                                        subCoverageVATAmountS = 0, // VAT Gián đoạn kinh doanh
                                        subCoveragePremiumAfterTaxS = 0; // Tổng phí phần II
                            var infoInsured = "";

                            // Có gián đoạn kinh doanh
                            if (hasCOVFBI)
                            {
                                var covFbiCount = certificate[j].RISKS.RISK.COVFBI.SUB_COVERAGE.Count();
                                /*     
                                   sum các giá trị Gián đoạn kinh doanh:
                                   subCoverageAnnualRateS Tỷ lệ phí Gián đoạn kinh doanh
                                   subCoveragePremiumS Phí Gián đoạn kinh doanh
                                   subCoverageVATAmountS VAT Gián đoạn kinh doanh
                                   subCoveragePremiumAfterTaxS Tổng phí phần II
                                */
                                for (int l = 0; l < covFbiCount; l++)
                                {
                                    var fbi = certificate[j].RISKS.RISK.COVFBI.SUB_COVERAGE[l];
                                    subCoverageAnnualRateS = fbi.SubCoverageAnnualRate;
                                    subCoveragePremiumS += fbi.SubCoveragePremium;
                                    subCoverageVATAmountS = (subCoveragePremiumS * 10 / 100);
                                    subCoveragePremiumAfterTaxS = (subCoveragePremiumS + subCoverageVATAmountS);
                                }
                            }
                            // nếu không có rủi ro khác, ngoài cháy nổ bắt buộc
                            if (hasCOVFSP != true && hasCOVFPA != true)
                            { }
                            // có mở rộng thêm Hỏa hoạn và các rủi ro đặc biệt hoặc Mọi rủi ro tài sản
                            else
                            {
                                var f = new SUB_COVERAGE();
                                // Mọi Rủi Ro Tài Sản
                                if (hasCOVFPA)
                                { }
                                // Bảo hiểm hỏa hạn và các rủi ro đặc biệt
                                else if (hasCOVFSP)
                                {
                                    var covFspCount = certificate[j].RISKS.RISK.COVFSP.SUB_COVERAGE.Count();
                                    // sum các giá trị covfsp
                                    for (int l = 0; l < covFspCount; l++)
                                    {
                                        f = certificate[j].RISKS.RISK.COVFSP.SUB_COVERAGE[l];
                                        subCoverageAnnualRateOther = f.SubCoverageAnnualRate;
                                        subCoveragePremiumOther += f.SubCoveragePremium;
                                        subCoverageVATOther = (subCoveragePremiumOther * 10 / 100);
                                    }
                                }
                                // Tỷ lệ phí bảo hiểm 
                                infoInsured += "Tỷ lệ phí bảo hiểm         : " + String.Format("{0:0.##}", (subCoverageAnnualRateOther) * 100).Replace(".", ",") + " % (chưa bao gồm VAT)" + newLine;
                                // Phí bảo hiểm
                                infoInsured += "Phí bảo hiểm                  : " + (subCoveragePremiumOther).CurrencyS(currencyCode) + " " + currencyCode + newLine;
                                // VAT
                                infoInsured += "VAT                               : " + (subCoverageVATOther).CurrencyS(currencyCode) + " " + currencyCode + newLine;
                                // Tổng phí
                                otherPremiumAfterTax = (subCoveragePremiumOther) + (subCoverageVATOther);
                                lstInsured.Add(new tb_Insured
                                {
                                    InfoInsured = infoInsured,
                                    STTRoman = NumberToString.IntToRoman(stt),
                                    SubCoveragePremiumAfterTax = otherPremiumAfterTax.CurrencyS(currencyCode), // Tổng phí phần I = (_f.SubCoveragePremium * 10 / 100).CurrencyS(_CurrencyCode), // VAT
                                                                                                               //// --Gián đoạn kinh doanh--
                                                                                                               // Tỷ lệ phí bảo hiểm
                                    SubCoverageAnnualRateS = (subCoverageAnnualRateS * 100).ToString(),
                                    // Phí bảo hiểm
                                    SubCoveragePremiumS = subCoveragePremiumS.CurrencyS(currencyCode),
                                    // VAT
                                    SubCoverageVATAmountS = (subCoverageVATAmountS * 10 / 100).CurrencyS(currencyCode),
                                    // Tổng phí phần II
                                    SubCoveragePremiumAfterTaxS = subCoveragePremiumAfterTaxS.CurrencyS(currencyCode),
                                    CurrencyCode = currencyCode,
                                    KeyChildTable = (j + 1).ToString()
                                });
                                stt++;
                            }

                        }
                        return lstInsured;
                    }
                    _dataModelView.addtable_lv2(DataDictionaryExtension.ConvertToDictionary(InsuredNF2S()), DataDictionaryExtension.ConvertToDictionary(InsuredSubNF2S()), "KeyChildTable", "~~~InsuredNF2S", "~~~InsuredSubNF2S", 0);
                    #endregion
                    break;
            }
            return _dataModelView;
        }

    }
}