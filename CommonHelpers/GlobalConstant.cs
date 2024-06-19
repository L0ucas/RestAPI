using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.Common
{
    public static class GlobalConstant
    {
        public const int CONST_SUCCESS_API_CODE = 200;
        public const string CONST_DEFAULT_CURRENCY = "IDR";
        public const string CONST_USD_CURRENCY = "USD";
        public const string CONST_SGD_CURRENCY = "SGD";
        public const string CONST_CONFIG_FILE_PATH = "R:\\";
        public const string CONST_SAP_MEMORY_LOW_DESC = "Memory low. Leave the transaction before taking a break!";
        public const string CONST_SAP_LIST_CONTAINS_NO_DATA_DESC = "List contains no data";
        public const string CONST_SAP_SCREENSHOT_DEFAULT_PNG_FILE_NAME = "RPA_SAP_ScreenShot_{0}.png";
        public const string CONST_NA = "N/A";
        public static readonly string[] CONST_MONTH_NM = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
        public static readonly string[] CONST_MONTH_ABB_ID = { "JAN", "FEB", "MAR", "APR", "MEI", "JUN", "JUL", "AUG", "SEP", "OKT", "NOV", "DEC" };
        public const int CONST_MAX_RETRY_CHECK_FILE = 5;
#if DEBUG
        public const string CONST_RPA_MASTER_CONFIG_FILE_PATH = @"C:\innersource\common\TMFramework\DotNet\TMHelper\TMHelper.Common\ConfigFile\RPAMaster\AppConfig.xml";
#else
        public const string CONST_RPA_MASTER_CONFIG_FILE_PATH = @"\\globalnet\aspiro\RPA\RPA_Working\Config\RPAMaster\AppConfig.xml";
#endif
        public const string CONST_CONFIG_KEY_DATABASE_IP = "DatabaseIP";
        public const string CONST_CONFIG_KEY_DATABASE_IP_IAP_IDN = "DatabaseIPIapIdn";
        public const string CONST_CONFIG_KEY_DATABASE_IP_IAP_CHN = "DatabaseIPIapChn";

        public static class TCode
        {
            public const string CONST_AUTO_MIRO_T_CODE = "ZM577N";
            public const string CONST_MIRO_T_CODE = "MIRO";
            public const string CONST_FBL1N_T_CODE = "FBL1N";
            public const string CONST_ZINV02A_T_CODE = "ZINV02A";
            public const string CONST_ME23N_T_CODE = "ME23N";
            public const string CONST_FB03_T_CODE = "FB03";
            public const string CONST_ZS411N_T_CODE = "ZS411N";
            public const string CONST_ME23_T_CODE = "ME23";
            public const string CONST_VA03_T_CODE = "VA03";
            public const string CONST_ME22_T_CODE = "ME22";
            public const string CONST_VA02_T_CODE = "VA02";
            public const string CONST_ZS796_T_CODE = "ZS796";
            public const string CONST_ZM88A_T_CODE = "ZM88A";
            public const string CONST_ZS460_T_CODE = "ZS460";
            public const string CONST_ZS450_T_CODE = "ZS450";
            public const string CONST_ZF251_T_CODE = "ZF251";
            public const string CONST_F_51_T_CODE = "F-51";
            public const string CONST_F_65_T_CODE = "F-65";
            public const string CONST_Z1MMIV02_T_CODE = "Z1MMIV02";
            public const string CONST_ZGESD14_T_CODE = "ZGESD14";
            public const string CONST_ZSTP00160_T_CODE = "ZSTP00160";
            public const string CONST_ZSTP00200_T_CODE = "ZSTP00200";
            public const string CONST_Z1SD460_T_CODE = "Z1SD460";
            public const string CONST_VF04_T_CODE = "VF04";
            public const string CONST_VF31_T_CODE = "VF31";
            public const string CONST_ZS787_T_CODE = "ZS787";
            public const string CONST_ZS13_T_CODE = "ZS13";
            public const string CONST_ZS908A_T_CODE = "ZS908A";
            public const string CONST_ZS52C_T_CODE = "ZS52C";
            public const string CONST_ZS849_T_CODE = "ZS849";
            public const string CONST_ZF88_T_CODE = "ZF88";
            public const string CONST_FS10N_T_CODE = "FS10N";
            public const string CONST_ZS332_T_CODE = "ZS332";
            public const string CONST_ZS52_T_CODE = "ZS52";
			public const string CONST_FK03_CODE = "FK03";
            public const string CONST_SE16N_T_CODE = "SE16N";
            public const string CONST_ZS815_T_CODE = "ZS815";
            public const string CONST_VF21_T_CODE = "VF21";
            public const string CONST_VF26_T_CODE = "VF26";
            public const string CONST_ZS962_T_CODE = "ZS962";
            public const string CONST_ZS967_T_CODE = "ZS967";
            public const string CONST_ZS925_T_CODE = "ZS925";
            public const string CONST_ZS970_T_CODE = "ZS970";
            public const string CONST_ZS971_T_CODE = "ZS971";
            public const string CONST_ZS871_T_CODE = "ZS871";
            public const string CONST_ZS984_T_CODE = "ZS984";
            public const string CONST_VF11_T_CODE = "VF11";
            public const string CONST_ZS986_T_CODE = "ZS986";
            public const string CONST_ZRTR00280_T_CODE = "ZRTR00280";
            public const string CONST_Z1FIEXP5_T_CODE = "Z1FIEXp5";
            public const string CONST_ZSSCDOCPR_T_CODE = "ZSSCDOCPR";
            public const string CONST_Z1FR_T_CODE = "Z1FR";
            public const string CONST_Z1FITLP5_T_CODE = "Z1FITLP5";
            public const string CONST_ZF124_T_CODE = "ZF124";
            public const string CONST_ZS1006_T_CODE = "ZS1006";
            public const string CONST_ZF343_T_CODE = "ZF343";
            public const string CONST_VF23_T_CODE = "VF23";
            public const string CONST_ZS454_T_CODE = "ZS454";
            public const string CONST_FBL3N_T_CODE = "FBL3N";
            public const string CONST_ZF_98_T_CODE = "ZF_98";
            public const string CONST_FB01_T_CODE = "FB01";
            public const string CONST_F_03_T_CODE = "F-03";
            public const string CONST_ZS410_T_CODE = "ZS410";
            public const string CONST_F110_T_CODE = "F110";
            public const string CONST_ZFI103N_T_CODE = "ZFI103N";
            public const string CONST_ZFI110_T_CODE = "ZFI110";
            public const string CONST_ZF402_T_CODE = "ZF402";
            public const string CONST_ZF378_T_CODE = "ZF378";
            public const string CONST_ZH34_T_CODE = "ZH34";
			public const string CONST_Z1HRTM010_A_CODE = "Z1HRTM010_A";
            public const string CONST_Z1HRTR04_CODE = "Z1HRTR04";
            public const string CONST_Z1HRTRT02_CODE = "Z1HRTRT02";
            public const string CONST_Z1HRTMT2011_CODE = "Z1HRTMT2011";
			public const string CONST_Z1HRPT004_CODE = "Z1HRPT004";
            public const string CONST_PT60_CODE = "PT60";
            public const string CONST_SMX_CODE = "SMX";
            public const string CONST_Z1HR_INFO_T_CODE = "Z1HR_INFO";
            public const string CONST_ZHR01_T_CODE = "ZHR01";
            public const string CONST_PA40_T_CODE = "PA40";
            public const string CONST_PA30_T_CODE = "PA30";
            public const string CONST_ZS1029_T_CODE = "ZS1029";
            public const string CONST_ZOTC01920_T_CODE = "ZOTC01920";
            public const string CONST_ZOTC01580_T_CODE = "ZOTC01580";
            public const string CONST_PA03_T_CODE = "PA03";
            public const string CONST_VF01_T_CODE = "VF01";
            public const string CONST_ZTR00128_T_CODE = "ZTR00128";
            public const string CONST_ZTR00011_T_CODE = "ZTR00011";
            public const string CONST_ZTRE00010 = "ZTRE00010";
            public const string CONST_ZTR00082 = "ZTR00082";
            public const string CONST_ZTR00033 = "ZTR00033";
            public const string CONST_ZTR00088 = "ZTR00088";
            public const string CONST_ZTR02006 = "ZTR02006";
            public const string CONST_ZF10_T_CODE = "ZF10";
            public const string CONST_ZHR01460_T_CODE = "ZHR01460";
            public const string CONST_ZHR01930_T_CODE = "ZHR01930";
            public const string CONST_ZHR02240_T_CODE = "ZHR02240";
            public const string CONST_ZHR03620_T_CODE = "ZHR03620";
            public const string CONST_PV34_T_CODE = "PV34";
            public const string CONST_PSV2_T_CODE = "PSV2";


        }

        public static class IAPWebService
        {
            public const string CONST_SUCCESS = "2";
            public const string CONST_FAILED = "3";
            public const string CONST_RESET = "9";
        }

        public static class TransactionCategory
        {
            public const string CONST_PO_FREIGHT_WAVE_ONE = "PO_Freight_W1a";
            public const string CONST_PO_FREIGHT_WAVE_TWO = "PO_Freight_W2a";
            public const string CONST_IMPORT_DUTY_AUTO_MIRO = "Import_Duty_Auto";
            public const string CONST_IMPORT_DUTY_MIRO = "Import_Duty_Manual";
            public const string CONST_ATL_AUTO_SO_PO = "ATL_Auto_SO_PO";
            public const string CONST_ATL_OCEAN_FREIGHT_MIRO = "ATL_Ocean_Freight_MIRO";
            public const string CONST_STANDARD_PO = "Standard_PO";
            public const string CONST_WASTE_PAPER = "Waste_Paper";
        }

        public static class CategoriesMethodName
        {
            public const string CONST_IAP_DATA_MASSAGE = "IAPDataMassage";
            public const string CONST_SAP_EXECUTION = "SAPExecution";
        }

        public static class AutoMiroFieldLength
        {
            public const int CONST_AUTO_MIRO_TEXT_FIELD_MAX_LENGTH = 50;
        }

        public static class TaxCode
        {
            public const string CONST_TAX_CODE_WITH_TAX_C1 = "C1";
            public const string CONST_TAX_CODE_WITHOUT_TAX_CN = "CN";
        }

        public static class ROCConstant
        {
            public const string CONST_ROC_FA_DEAL = "FA";
            public const string CONST_ROC_HR_DEAL = "HR";
            public const string CONST_ROC_SUCCESS_STATUS = "Success";
            public const string CONST_ROC_FAILED_STATUS = "Fail";
        }

        public static class PaymentTerm
        {
            public const string CONST_PAYMENT_TERM_OTHERS = "Others";
            public const string CONST_PAYMENT_TERM_N001 = "N001";
            public const string CONST_PAYMENT_TERM_N010 = "N010";
            public const string CONST_PAYMENT_TERM_N014 = "N014";
            public const string CONST_PAYMENT_TERM_N030 = "N030";
            public const string CONST_PAYMENT_TERM_COD = "COD";
        }
        public static class SAPTransactionType
        {
            public const string CONST_SAP_TRANSACTION_TYPE_INVOICE = "Invoice";
        }
        public static class ProductAssignment
        {
            public const string CONST_PRODUCT_ASSIGNMENT_LIMESTONE = "Limestone";
            public const string CONST_PRODUCT_ASSIGNMENT_WASTEPAPER = "Wastepaper";
            public const string CONST_PRODUCT_ASSIGNMENT_COAL = "Coal";
            public const string CONST_PRODUCT_ASSIGNMENT_PULP_TRUCKING = "Pulp – Trucking";
            public const string CONST_PRODUCT_ASSIGNMENT_PULP_HANDLING = "Pulp – Handling";
            public const string CONST_PRODUCT_ASSIGNMENT_PULP_FREIGHT = "Pulp – Freight";
        }
        public static class ProductAssignmentText
        {
            public const string CONST_PRODUCT_ASSIGNMENT_TEXT_LIMESTONE = "B. Ongkos Angkut Lime Stone";
            public const string CONST_PRODUCT_ASSIGNMENT_TEXT_WASTEPAPER = "B. Ongkos Angkut Waste Paper";
            public const string CONST_PRODUCT_ASSIGNMENT_TEXT_COAL = "B. Ongkos Angkut Batu Bara";
            public const string CONST_PRODUCT_ASSIGNMENT_TEXT_PULP_TRUCKING = "B. Trucking Pulp";
            public const string CONST_PRODUCT_ASSIGNMENT_TEXT_PULP_HANDLING = "B. Handling Pulp";
            public const string CONST_PRODUCT_ASSIGNMENT_TEXT_PULP_FREIGHT = "B. Freight Pulp";
        }
        public static class MiroPosting
        {
            public const string CONST_MIRO_POSTING_DETAIL_DOC_TYPE_RE = "RE";
            public const string CONST_MIRO_POSTING_PO_REF_PO_TYPE = "1";
            public const string CONST_MIRO_POSTING_PO_REF_PO_TYPE_DELIVERY_NOTE = "2";
            public const string CONST_MIRO_POSTING_PO_REF_PO_TYPE_GOODS_SERVICE_ITEMS_PLANNED_DELIVERY_COSTS = "3";
            public const string CONST_MIRO_POSTING_PO_REF_FS_ITEM_TYPE_PLANNED_DELIVERY_COST = "2";
            public const string CONST_MIRO_POSTING_PO_REF_FS_ITEM_TYPE_GOODS_SERVICE_ITEMS = "1";
            public const string CONST_MIRO_POSTING_PO_REF_LAYOUT_ACCT_HQ_INFO = "Z_7_6310_ACCT HQ";
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ACCT_HQ_INFO_COL_IDX_PO_TEXT = 6;
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ACCT_HQ_INFO_COL_IDX_Quantity = 3;
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ACCT_HQ_INFO_COL_IDX_AMOUNT = 2;
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ACCT_HQ_INFO_COL_IDX_TAX_CODE = 9;
            public const string CONST_MIRO_POSTING_PO_REF_LAYOUT_PO_HISTORY = "3_6310";
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_PO_HISTORY_COL_IDX_PO_TEXT = 31;
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_PO_HISTORY_COL_IDX_Quantity = 5;
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_PO_HISTORY_COL_IDX_AMOUNT = 7;
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_PO_HISTORY_COL_IDX_TAX_CODE = 8;
            public const string CONST_MIRO_POSTING_PO_REF_LAYOUT_ALL_INFO = "7_6310";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_PO_TEXT = "Short Text";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_PO_NO = "Purchase Order";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_PO_ITEM_NO = "Item";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_QUANTITY = "Quantity";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_AMOUNT = "Amount";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_TAX_CODE = "Tax code";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_INV_UNIT_PRICE = "Price Unit";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_INV_QUANTITY = "Quantity Ordered";             //PO Quantity
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_QUANTITY_RECEIVED = "Quantity Received";             //Received
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_QUANTITY_SETTLED = "Quantity invoiced";             //Settle
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_SUB_CR_DB = "Subseq. debit/credit";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_MATERIAL = "Material";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_REFERENCE_DOC = "Reference Document";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_DELIVERY_NOTE = "Delivery Note";
            public const string CONST_MIRO_POSTING_PO_REF_COL_NM_TEXT = "Text";
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_PO_TEXT = "txtDRSEG-TXZ01[{0},{1}]";                 // PO Text
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_PO_NO = "txtDRSEG-EBELN[{0},{1}]";                   // PO Number
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_PO_ITEM_NO = "txtDRSEG-EBELP[{0},{1}]";              // PO Item Number
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_QUANTITY = "txtDRSEG-MENGE[{0},{1}]";                // Quantity
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_AMOUNT = "txtDRSEG-WRBTR[{0},{1}]";                  // Amount
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_TAX_CODE = "cmbDRSEG-MWSKZ[{0},{1}]";                // Tax Code
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_INV_UNIT_PRICE = "txtDRSEG-PEINH[{0},{1}]";          // Unit Price
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_INV_QUANTITY = "txtDRSEG-BSMNG[{0},{1}]";            // PO Quantity
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_QUANTITY_RECEIVED = "txtDRSEG-WEMNG[{0},{1}]";       // Received
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_QUANTITY_SETTLED = "txtDRSEG-REMNG[{0},{1}]";        // Settle
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_SUB_CR_DB = "chkDRSEG-TBTKZ[{0},{1}]";               // Subsequence Debit 
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_MATERIAL = "ctxtDRSEG-MATNR[{0},{1}]";               // PO Material
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_REFERENCE_DOC = "txtDRSEG-LFBNR[{0},{1}]";       //Reference Doc
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_DELIVERY_NOTE = "txtDRSEG-XBLNR[{0},{1}]";       //Delivery Note
            public const string CONST_MIRO_POSTING_PO_REF_FIELD_ID_TEXT = "txtDRSEG-SGTXT[{0},{1}]";
            public const string CONST_MIRO_POSTING_GL_ACCT_LAYOUT_STANDARD_4 = "Standard 4";
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_GL_ACCT = "ctxtACGL_ITEM-HKONT[{0},{1}]";       //G/L acct
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_DC = "cmbACGL_ITEM-SHKZG[{0},{1}]";             //D/C
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_AMOUNT = "txtACGL_ITEM-WRBTR[{0},{1}]";         //Amount in doc.curr
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_TAX_CODE = "ctxtACGL_ITEM-MWSKZ[{0},{1}]";      //Tax Code
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_ASSIGNMENT = "txtACGL_ITEM-ZUONR[{0},{1}]";     //Assignment
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_TEXT = "ctxtACGL_ITEM-SGTXT[{0},{1}]";          //Text
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_COST_CENTER = "ctxtACGL_ITEM-KOSTL[{0},{1}]";   //Cost Center
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_FUND_CENTER = "ctxtACGL_ITEM-FISTL[{0},{1}]";   //Fund Center
            public const string CONST_MIRO_POSTING_GL_ACCT_FIELD_ID_COMPANY_CODE = "ctxtACGL_ITEM-BUKRS[{0},{1}]";
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_GL_ACCT = "G/L Account";              //G/L acct
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_DC = "Debit/Credit Ind.";             //D/C
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_AMOUNT = "Amount";                    //Amount in doc.curr
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_TAX_CODE = "Tax code";                //Tax Code
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_ASSIGNMENT = "Assignment";            //Assignment
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_TEXT = "Text";                        //Text
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_COST_CENTER = "Cost Center";          //Cost Center
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_FUND_CENTER = "Funds Center";          //Cost Center
            public const string CONST_MIRO_POSTING_GL_ACCT_COL_NM_COMPANY_CODE = "Company Code";         //Company Code
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ALL_INFO_COL_IDX_PO_TEXT = 7;             //txtDRSEG-TXZ01[7,0]
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ALL_INFO_COL_IDX_PO_NO = 5;               //txtDRSEG-EBELN[5,0]
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ALL_INFO_COL_IDX_Quantity = 2;            //txtDRSEG-MENGE[2,0]
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ALL_INFO_COL_IDX_AMOUNT = 1;              //txtDRSEG-WRBTR[1,0]
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ALL_INFO_COL_IDX_TAX_CODE = 8;            //cmbDRSEG-MWSKZ[8,0]
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ALL_INFO_COL_IDX_INV_UNIT_PRICE = 56;     //txtDRSEG-PEINH[56, 0]
            public const int CONST_MIRO_POSTING_PO_REF_LAYOUT_ALL_INFO_COL_IDX_INV_QUANTITY = 49;       //txtDRSEG-BSMNG[49,0]
            public const string CONST_MIRO_POSTING_PO_REF_FREIGHT_Q_3RD_PARTY = "Freight/Q-3rd party";
            public const string CONST_MIRO_POSTING_PO_REF_FREIGHT_3RD_PARTY = "Freight%-3rd party";
            public const string CONST_MIRO_POSTING_GL_ACCT_BIAYA_BONGKAR_GL = "6480045";
            public const string CONST_MIRO_POSTING_GL_ACCT_PULP_GL = "1420105";
            public const string CONST_MIRO_POSTING_GL_ACCT_PPH = "1420102";
            public const string CONST_MIRO_POSTING_GL_ACCT_PPN = "1420105";
            public const string CONST_MIRO_POSTING_GL_ACCT_PPH_TEXT = "T22-{0}";
            public const string CONST_MIRO_POSTING_GL_ACCT_PPN_TEXT = "VAT-{0}";
            public const string CONST_MIRO_POSTING_GL_ACCT_HQ_COST_CENTER = "8368308000";
            public const string CONST_MIRO_POSTING_GL_ACCT_TEXT_BIAYA_BONGKAR = "Biaya Bongkar";
            public const string CONST_MIRO_POSTING_GL_ACCT_TEXT_LIMESTONE = "Limestone";
            public const string CONST_MIRO_POSTING_GL_ACCT_TEXT_DEBIT = "Debit";
            public const string CONST_MIRO_POSTING_GL_ACCT_TEXT_CREDIT = "Credit";
            public const string CONST_MIRO_POSTING_INV_ADY_INSERTED = "Check if invoice already entered under accounting doc. no.";
            public static readonly string[] CONST_MIRO_POSTING_IMPORT_DUTY_PO_TEXT = { "Import Duty", "Financial Charges", "Freight Charges", "Freight" };
            public const string CONST_MIRO_POSTING_TEXT_PERMITTED_PAYEE = "PERMITTED PAYEE";
            public const string CONST_MIRO_POSTING_TEXT_PPH23 = "PPH23";
            public const string CONST_MIRO_POSTING_GL_ACCT_PPH22_ASSIGNMENT = "KRW";
            public const string CONST_MIRO_POSTING_GL_ACCT_PPH22_COMPANY_CODE = "8006";
            public const string CONST_MIRO_POSTING_INVALID_REF_MESSAGE = "Invalid Reference- Check Rcpt Note->";
        }

        public static class MiroPostingException
        {
            public const string CONST_MIRO_POSTING_EXCEPTION_SAP_FAILED = "MIRO - Posting Failed. ";

        }

        public static class AutoMiroGLAccountNo
        {
            public const string CONST_PPN_GL_ACCOUNT_NO = "1420105";
            public const string CONST_PPH_GL_ACCOUNT_NO = "1420102";
        }

        public static class ImportDutyTaxCode
        {
            public const string CONST_PPN_TAX_CODE = "C2";
            public const string CONST_PPH_TAX_CODE = "A3";
            public const string CONST_C2_TAX_CODE = "C2";
            public const string CONST_CN_TAX_CODE = "CN";
        }

        public static class ImportDutyDummyVendor
        {
            public const string CONST_TK100IMPOR_VENDOR = "TK100IMPOR";
            public const string CONST_IT01IMPORT_VENDOR = "IT01IMPORT";
            // IKP 
            public const string CONST_IK100IMPOR_VENDOR = "IK100IMPOR";
            public const string CONST_LP100IMPOR_VENDOR = "LP100IMPOR";
            public const string CONST_IMPORT_VENDOR = "IMPORT";
            public const string CONST_OKI10IMPOR_VENDOR = "OKI10IMPOR";
            // IKP
            // IKQ
            //public const string CONST_IK100IMPOR_VENDOR = "17329";
            //public const string CONST_LP100IMPOR_VENDOR = "15975";
            //public const string CONST_IMPORT_VENDOR = "86777";
            //public const string CONST_OKI10IMPOR_VENDOR = "92641";
            // IKQ
        }

        public static class InvoiceType
        {
            public const string CONST_EMKL_INVOICE_TYPE = "EMKL";
            public const string CONST_REIMBURSEMENT_INVOICE_TYPE = "Reimbursement";
            public const string CONST_THC_INVOICE_TYPE = "THC";
            public const string CONST_COURIER_INVOICE_TYPE = "Courier";
        }

        public static class POText
        {
            public const string CONST_IMPORT_DUTY_PO_TEXT = "Import duty";
            public const string CONST_FREIGHT_Q_PO_TEXT = "FREIGHT/Q";
            public const string CONST_FREIGHT_PERCENT_PO_TEXT = "FREIGHT-%";
            public const string CONST_FREIGHT_PERCENT_PO_TEXT_SECOND_STYLE = "Freight%";
            public const string CONST_FREIGHT_V_PO_TEXT = "FREIGHTV";
            public const string CONST_FIN_CHARGES = "Fin.charges";
            public const string CONST_ORIGINAL_FREIGHT_PO_TEXT_ABBR = "OF";
            public const string CONST_ORIGINAL_FREIGHT_PO_TEXT = "Original Freight";
            public const string CONST_FREIGHT_FORWARDER_FEE_PO_TEXT = "Frght Forwarder Fee";
        }

        public static class ImportDutyTextFieldValue
        {
            public const string CONST_BEA_MASUK_TEXT_FIELD_VALUE = "BEA MASUK";
            public const string CONST_PPN_TEXT_FIELD_VALUE = "PPN IMPOR";
            public const string CONST_PPH_TEXT_FIELD_VALUE = "PPH IMPOR";
        }

        public static class ATLPosting
        {
            public const string CONST_ATL_SO_PO_STATUS_SUCCESS = "Success";
            public const string CONST_ATL_SO_PO_STATUS_FAILED = "Failed";
            public const string CONST_ATL_OCEAN_FREIGHT_MIRO_STATUS_SUCCESS = "Success";
            public const string CONST_ATL_OCEAN_FREIGHT_MIRO_STATUS_FAILED = "Failed";
            public const double CONST_ATL_OCEAN_FREIGHT_MIRO_PPH23_RATE = 0.02;
            public const string CONST_ATL_SO_PO_SYNC_SALES_ORG = "2396";

        }

        public static class VA03TextType
        {
            public const string CONST_ATL_CONTRACT_NO_TEXT = "ATL Contract No";
            public const string CONST_ATL_CONTRACT_NO_NODE_ID = "ZS25";
        }

        public static class ATLConditionType
        {
            public const string CONST_ORIGINAL_FREIGHT = "ZRA5";
            public const string CONST_FREIGHT_FWDER = "ZRA6";
        }

        public static class RPATaskStatus
        {
            public const int CONST_RPA_TASK_SUCCESS = 0;
            public const int CONST_RPA_TASK_FAIL = 1;
            public const int CONST_RPA_TASK_SAP_TIMEOUT = -99;
        }

        public static class RPAMasterConfig
        {
            // Dev
            //public const string CONST_CONFIG_KEY_EMAIL_PWD = "RPAEmailPassword_Dev";
            // Prod
            public const string CONST_CONFIG_KEY_EMAIL_PWD = "RPAEmailPassword";
            public const string CONST_CONFIG_KEY_EXCHANGE_USERNAME = "RPAExchangeUsername";
            public const string CONST_CONFIG_KEY_EXCHANGE_PWD = "RPAExchangePwd";
            public const string CONST_CONFIG_KEY_EXCHANGE_SVC_URL = "RPAExchangeServiceURL";
            public const string CONST_CONFIG_KEY_EXCHANGE_FROM_EMAIL_ADDR = "RPAExchangeServiceFromEmailAddr";
            public const string CONST_COL_NM_CONFIG_KEY = "ConfigKey";
            public const string CONST_COL_NM_CONFIG_VALUE = "ConfigValue";
            public const string CONST_IBM_SERVER_DOMINO = "DominoT/TEST";
            public const string CONST_CONFIG_KEY_RPA_MS_GRAPH_API_TENANT_ID = "RPAMsGraphApiTenantId";
            public const string CONST_CONFIG_KEY_RPA_MS_GRAPH_API_CLIENT_SECRET = "RPAMsGraphApiClientSecret";
            public const string CONST_CONFIG_KEY_RPA_MS_GRAPH_API_CLIENT_ID = "RPAMsGraphApiClientId";
            public const string CONST_CONFIG_KEY_RPA_MS_GRAPH_API_USERNAME = "RPAMsGraphApiUsername";
        }

        public static class StoreProcedureName
        {
            public const string CONST_USP_RPA_CALENDAR_SEL = "UspRPACalendarSel";
            public const string CONST_USP_RPA_TOOL_CONFIG_SEL = "UspRPAToolConfigSel";
            public const string CONST_USP_RPA_PARAMETER_SEL = "UspRPAParameterSel";
            public const string CONST_USP_RPA_CREDENTIAL_SEL = "UspRPACredentialSel";
            public const string CONST_USP_RPA_PARAMETER_UPD = "UspRPAParameterUpd";
            public const string CONST_USP_RPA_CREDENTIAL_UPD = "UspRPACredentialUpd";
            public const string CONST_USP_RPA_EMAIL_TEMPLATE_SEL = "UspRPAEmailTemplateGrpSel";
        }

        public static class TissueTypeCode
        {
            public const string CONST_TISSUE_TRADING = "TT";
            public const string CONST_TISSUE_MILL = "TM";
        }

        public static class CNPayPostProcessStats
        {
            public const string CONST_STATS_INITIAL = "1";
            public const string CONST_STATS_VALIDATED = "2";
            public const string CONST_STATS_FAILED_REQUIRE_HUMAN_INTERVENTION = "3";
            public const string CONST_STATS_FAILED_SYSTEM_ERROR = "4";
            public const string CONST_STATS_COMPLETED = "5";
        }

        public static class CNPayPostPAPymtTypeCode
        {
            public const string PA_PAYMENT_TYPE_BANK_TRANSFER = "1";
            public const string PA_PAYMENT_TYPE_GUARANTEE_PAYMENT = "2";

        }

        public static class SAPInstance
        {
            public enum Indonesia
            {
                IKP = 0,
                P01 = 1,
                HR1 = 2,
                P02 = 3,
                S4 = 4
            }

            public enum China
            {
                NAPlus = 0,
                APlus = 1
            }
        }
    }
}
