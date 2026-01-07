using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests;
using Shouldly;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.IE.Tests
{
    public class CustomTemplateExportTest:TestBase
    {

        [Fact(DisplayName = "自定义模板导出测试")]
        public async Task TemplateExport()
        {
            var json = $"{{\r\n   \"CustomerEmail\":[\r\n      \"sales@testdata.com\"\r\n   ],\r\n   \"CustomerName\":\"\",\r\n   \"CustomerEmailString\":\"sales@testdata.com\",\r\n   \"CsEmail\":[\r\n      \"my@testdata.com\"\r\n   ],\r\n   \"CsEmailString\":\"my@testdata.com\",\r\n   \"CsDisplayEmailString\":\"my@testdata.com\",\r\n   \"CcEmail\":[\r\n      \r\n   ],\r\n   \"CcEmailString\":null,\r\n   \"GroupOrgKey\":\"ORG11111112222222\",\r\n   \"GroupKey\":\"1111111\\u00262222222\",\r\n   \"ORG_CODE\":\"ORG\",\r\n   \"SOLD_TO_NUMBER\":\"1111111\",\r\n   \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n   \"BILL_TO_NUMBER\":\"2222222\",\r\n   \"BILL_TO_CUSTOMER\":\"customer\",\r\n   \"AmtSummary\":42475.811400,\r\n   \"TaxSummary\":5521.89,\r\n   \"AmtTaxSummary\":47997.701400,\r\n   \"CartonSummary\":16,\r\n   \"WeightSummary\":85.600,\r\n   \"CbmSummary\":0.18747826,\r\n   \"Data\":[\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"1\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168939/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":2638,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":1284.917040,\r\n         \"TAX_AMOUNT\":167.04,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"1284.917040\",\r\n         \"AMT_TAX\":1451.957040,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":2.550,\r\n         \"CBM\":0.00449868,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref\",\r\n         \"GOES_ORDER_AND_LINE\":\"order ref line\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"2\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168861/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":6697,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":3261.974760,\r\n         \"TAX_AMOUNT\":424.06,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"3261.974760\",\r\n         \"AMT_TAX\":3686.034760,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":6.420,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329726522-2\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4M_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"3\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168860/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":5152,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2509.436160,\r\n         \"TAX_AMOUNT\":326.23,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2509.436160\",\r\n         \"AMT_TAX\":2835.666160,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":5.040,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329726522-3\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4N_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"4\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168859/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":5153,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2509.923240,\r\n         \"TAX_AMOUNT\":326.29,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2509.923240\",\r\n         \"AMT_TAX\":2836.213240,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":5.040,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329726522-4\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4P_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"5\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168858/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":6697,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":3261.974760,\r\n         \"TAX_AMOUNT\":424.06,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"3261.974760\",\r\n         \"AMT_TAX\":3686.034760,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":6.435,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329726522-5\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4R_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"6\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168857/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":5152,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2509.436160,\r\n         \"TAX_AMOUNT\":326.23,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2509.436160\",\r\n         \"AMT_TAX\":2835.666160,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":5.020,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329726522-6\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4T_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"7\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168855/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":5153,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2509.923240,\r\n         \"TAX_AMOUNT\":326.29,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2509.923240\",\r\n         \"AMT_TAX\":2836.213240,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":5.120,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329726522-7\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4U_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"8\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168844/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":1582,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":770.560560,\r\n         \"TAX_AMOUNT\":100.17,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"770.560560\",\r\n         \"AMT_TAX\":870.730560,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":5.635,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329726522-8\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4V_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"9\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168843/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":5201,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2533.303080,\r\n         \"TAX_AMOUNT\":329.33,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2533.303080\",\r\n         \"AMT_TAX\":2862.633080,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":5.125,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329726522-9\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4W_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"10\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168841/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":3152,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":1535.276160,\r\n         \"TAX_AMOUNT\":199.59,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"1535.276160\",\r\n         \"AMT_TAX\":1734.866160,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":8.580,\r\n         \"CBM\":0.02361419,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref0\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F4X_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"11\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168840/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":4682,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2280.508560,\r\n         \"TAX_AMOUNT\":296.47,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2280.508560\",\r\n         \"AMT_TAX\":2576.978560,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":4.600,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref1\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F50_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"12\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168839/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":2626,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":1279.072080,\r\n         \"TAX_AMOUNT\":166.28,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"1279.072080\",\r\n         \"AMT_TAX\":1445.352080,\r\n         \"CartonCount\":0,\r\n         \"SOL_GW\":0,\r\n         \"CBM\":0,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref2\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F51_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"13\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168838/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":2942,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":1432.989360,\r\n         \"TAX_AMOUNT\":186.29,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"1432.989360\",\r\n         \"AMT_TAX\":1619.279360,\r\n         \"CartonCount\":0,\r\n         \"SOL_GW\":0,\r\n         \"CBM\":0,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref3\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F52_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"14\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168837/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":2731,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":1330.215480,\r\n         \"TAX_AMOUNT\":172.93,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"1330.215480\",\r\n         \"AMT_TAX\":1503.145480,\r\n         \"CartonCount\":0,\r\n         \"SOL_GW\":0,\r\n         \"CBM\":0,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref4\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F53_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"15\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168836/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":5202,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2533.790160,\r\n         \"TAX_AMOUNT\":329.39,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2533.790160\",\r\n         \"AMT_TAX\":2863.180160,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":5.075,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref5\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F54_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"16\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168834/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":2942,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":1432.989360,\r\n         \"TAX_AMOUNT\":186.29,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"1432.989360\",\r\n         \"AMT_TAX\":1619.279360,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":7.205,\r\n         \"CBM\":0.02361419,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref6\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F55_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"17\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168833/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":1472,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":716.981760,\r\n         \"TAX_AMOUNT\":93.21,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"716.981760\",\r\n         \"AMT_TAX\":810.191760,\r\n         \"CartonCount\":0,\r\n         \"SOL_GW\":0,\r\n         \"CBM\":0,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref7\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F56_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"18\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168831/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":4681,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2280.021480,\r\n         \"TAX_AMOUNT\":296.40,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2280.021480\",\r\n         \"AMT_TAX\":2576.421480,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":4.680,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref8\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F57_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"19\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800168830/457\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":4727,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2302.427160,\r\n         \"TAX_AMOUNT\":299.32,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2302.427160\",\r\n         \"AMT_TAX\":2601.747160,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":4.710,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"order ref9\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F58_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"20\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800170912/18722\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":4417,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2151.432360,\r\n         \"TAX_AMOUNT\":279.69,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2151.432360\",\r\n         \"AMT_TAX\":2431.122360,\r\n         \"CartonCount\":1,\r\n         \"SOL_GW\":4.365,\r\n         \"CBM\":0.01044240,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329727617-1\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F59_0001\"\r\n      }},\r\n      {{\r\n         \"ORDER_NUMBER\":\"order number\",\r\n         \"LINE_NUMBER\":\"21\",\r\n         \"ORG_CODE\":\"ORG\",\r\n         \"SOLD_TO_NUMBER\":\"1111111\",\r\n         \"SOLD_TO_CUSTOMER\":\"customer name\",\r\n         \"SET_NAME\":\"1\",\r\n         \"ORDERED_DATE\":\"12/29/2025 12:00:00 AM\",\r\n         \"ORDERED_DATE_Str\":\"2025-Dec-29\",\r\n         \"CUST_PO_NUMBER\":\"6800170911/18722\",\r\n         \"ITEM\":\"itemname2\",\r\n         \"ORDERED_ITEM\":\"itemname\",\r\n         \"QTY\":4206,\r\n         \"PRICE\":\"0.487080\",\r\n         \"AMT_Decimal\":2048.658480,\r\n         \"TAX_AMOUNT\":266.33,\r\n         \"CURRENCY_CODE\":\"CNY\",\r\n         \"AMT\":\"2048.658480\",\r\n         \"AMT_TAX\":2314.988480,\r\n         \"CartonCount\":0,\r\n         \"SOL_GW\":0,\r\n         \"CBM\":0,\r\n         \"BILL_TO_CUSTOMER\":\"customer\",\r\n         \"BILL_TO_NUMBER\":\"2222222\",\r\n         \"SHIP_TO_CUSTOMER\":\"ship to customer\",\r\n         \"CS\":\"my@testdata.com\",\r\n         \"WEB_ORDER_REF\":\"AD-329727617-2\",\r\n         \"GOES_ORDER_AND_LINE\":\"NT8F5A_0001\"\r\n      }}\r\n   ]\r\n}}";

            var obj = JsonSerializer.Deserialize<CustomerGroupDataVO>(json);

            var exporter = new ExcelExporter();
            var temp = Path.Combine(".","TestFiles", "ExportTemplates", "CustomExportTemplate.xlsx");
            var file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{Guid.NewGuid()}.xlsx");
            await exporter.ExportByTemplate(file, obj, temp);
            // 2.8+版本生成的excel有大量重复的模板内容:{{Table>>}} ... , 文件超过30M,2.7+版本生成的文件只有不到10kb
            var file_size = (new FileInfo(file)).Length;

            file_size.ShouldBeLessThan(10*1024);
        }
    }

    [Exporter(AutoFitAllColumn = true)]
    public class CustomerGroupDataVO
    {

        [ExporterHeader(IsIgnore = true)]
        public List<string?> CustomerEmail { get; set; } = new List<string?>();

        [ExporterHeader(IsIgnore = true)]
        public string? CustomerName { get; set; }

        [ExporterHeader(IsIgnore = true)]
        public string? CustomerEmailString => string.Join(",", CustomerEmail);

        [ExporterHeader(IsIgnore = true)]
        public List<string?> CsEmail { get; set; } = new List<string?>();

        [ExporterHeader(IsIgnore = true)]
        public string? CsEmailString => string.Join(",", CsEmail);

        [ExporterHeader(IsIgnore = true)]
        public string? CsDisplayEmailString { get; set; }

        [ExporterHeader(IsIgnore = true)]
        public List<string?> CcEmail { get; set; } = new List<string?>();

        [ExporterHeader(IsIgnore = true)]
        public string? CcEmailString => string.Join(",", CcEmail);

        [ExporterHeader(IsIgnore = true)]
        public string? GroupOrgKey => ORG_CODE + SOLD_TO_NUMBER + BILL_TO_NUMBER;

        [ExporterHeader(IsIgnore = true)]
        public string? GroupKey => SOLD_TO_NUMBER + "&" + BILL_TO_NUMBER;

        public string? ORG_CODE { get; set; }
        public string? SOLD_TO_NUMBER { get; set; }
        public string? SOLD_TO_CUSTOMER { get; set; }
        public string? BILL_TO_NUMBER { get; set; }
        public string? BILL_TO_CUSTOMER { get; set; }

        public decimal? AmtSummary => Data.Sum(a => a.AMT_Decimal);

        public decimal? TaxSummary => Data.Sum(a => a.TAX_AMOUNT);

        public decimal? AmtTaxSummary => Data.Sum(a => a.AMT_TAX);

        public int? CartonSummary => Data.Sum(a => a.CartonCount);

        public decimal? WeightSummary => Data.Sum(a => a.SOL_GW);

        public decimal? CbmSummary => Data.Sum(a => a.CBM);

        public List<InventoryCustomerDataVO> Data { get; set; } = new List<InventoryCustomerDataVO>();

    }

    [Exporter(AutoFitAllColumn = true)]
    public class InventoryCustomerDataVO
    {
        public string? ORDER_NUMBER { get; set; }

        public string? LINE_NUMBER { get; set; }
        public string? ORG_CODE { get; set; }
        public string? SOLD_TO_NUMBER { get; set; }
        public string? SOLD_TO_CUSTOMER { get; set; }

        public string? SET_NAME { get; set; }

        [ExporterHeader(IsIgnore = true)]
        public string? ORDERED_DATE { get; set; }

        public string? ORDERED_DATE_Str => DateTime.TryParse(ORDERED_DATE, out var val) ? val.ToString("yyyy-MMM-dd") : null;

        public string? CUST_PO_NUMBER { get; set; }

        public string? ITEM { get; set; }

        public string? ORDERED_ITEM { get; set; }


        public decimal? QTY { get; set; }

        public string? PRICE { get; set; }

        [ExporterHeader(IsIgnore = true)]
        public decimal? AMT_Decimal => decimal.TryParse(AMT, out var amt) ? amt : 0;

        public decimal? TAX_AMOUNT { get; set; }


        public string? CURRENCY_CODE { get; set; }

        public string? AMT { get; set; }


        public decimal? AMT_TAX => AMT_Decimal + TAX_AMOUNT;

        [ExporterHeader(DisplayName = "No.of Cartons")]
        public int CartonCount { get; set; }

        [ExporterHeader(DisplayName = "Weight(kg)")]
        public decimal? SOL_GW { get; set; }

        public decimal? CBM { get; set; }

        public string? BILL_TO_CUSTOMER { get; set; }

        public string? BILL_TO_NUMBER { get; set; }

        public string? SHIP_TO_CUSTOMER { get; set; }

        public string? CS { get; set; }

        public string? WEB_ORDER_REF { get; set; }

        public string? GOES_ORDER_AND_LINE { get; set; }
    }
}
