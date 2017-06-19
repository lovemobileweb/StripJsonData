Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Web.Script.Services
Imports Newtonsoft.Json
Imports System.IO

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class WebService
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function HelloWorld() As String
        Return "Hello World"
    End Function

    <WebMethod(ResponseFormat.Json)>
    <System.Web.Script.Services.ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
    Public Sub addrexxValidationCheck2JSON(location As String, callback As String)
        'Dim response As String = (XxerddaWeather.addressValidate2(location, callback))
        'Dim callbackValue As String = HttpContext.Current.Request.QueryString("callback")
        HttpContext.Current.Response.ContentType = "application/x-javascript"
        Dim response As String = "{
  ""authenticationResultCode"": ""ValidCredentials"",
  ""brandLogoUri"": ""http://dev.virtualearth.net/Branding/logo_powered_by.png"",
  ""copyright"": ""Copyright © 2017 Microsoft and its suppliers. All rights reserved. This API cannot be accessed and the content and any results may not be used, reproduced or transmitted in any manner without express written permission from Microsoft Corporation."",
  ""resourceSets"": [
    {
      ""estimatedTotal"": 4,
      ""resources"": [
        {
          ""__type"": ""Location:http://schemas.microsoft.com/search/local/ws/rest/v1"",
          ""bbox"": [
            41.1428913351492,
            -73.4990319661673,
            41.1506167702906,
            -73.4853538678185
          ],
          ""name"": ""123 Main St, New Canaan, CT 06840"",
          ""point"": {
            ""type"": ""Point"",
            ""coordinates"": [
              41.1467540527199,
              -73.4921929169929
            ]
          },
          ""address"": {
            ""addressLine"": ""123 Main St"",
            ""adminDistrict"": ""CT"",
            ""adminDistrict2"": ""Fairfield"",
            ""countryRegion"": ""United States"",
            ""formattedAddress"": ""123 Main St, New Canaan, CT 06840"",
            ""locality"": ""New Canaan"",
            ""postalCode"": ""06840""
          },
          ""confidence"": ""Medium"",
          ""entityType"": ""Address"",
          ""geocodePoints"": [
            {
              ""type"": ""Point"",
              ""coordinates"": [
                41.1467540527199,
                -73.4921929169929
              ],
              ""calculationMethod"": ""InterpolationOffset"",
              ""usageTypes"": [
                ""Display""
              ]
            },
            {
              ""type"": ""Point"",
              ""coordinates"": [
                41.1467650000053,
                -73.4921349999439
              ],
              ""calculationMethod"": ""Interpolation"",
              ""usageTypes"": [
                ""Route""
              ]
            }
          ],
          ""matchCodes"": [
            ""Ambiguous""
          ]
        },
        {
          ""__type"": ""Location:http://schemas.microsoft.com/search/local/ws/rest/v1"",
          ""bbox"": [
            42.5852272824293,
            -74.3286250955868,
            42.5929527175707,
            -74.3146349044132
          ],
          ""name"": ""123 Main St, Middleburgh, NY 12122"",
          ""point"": {
            ""type"": ""Point"",
            ""coordinates"": [
              42.58909,
              -74.32163
            ]
          },
          ""address"": {
            ""addressLine"": ""123 Main St"",
            ""adminDistrict"": ""NY"",
            ""adminDistrict2"": ""Schoharie"",
            ""countryRegion"": ""United States"",
            ""formattedAddress"": ""123 Main St, Middleburgh, NY 12122"",
            ""locality"": ""Middleburgh"",
            ""postalCode"": ""12122""
          },
          ""confidence"": ""Medium"",
          ""entityType"": ""Address"",
          ""geocodePoints"": [
            {
              ""type"": ""Point"",
              ""coordinates"": [
                42.58909,
                -74.32163
              ],
              ""calculationMethod"": ""Rooftop"",
              ""usageTypes"": [
                ""Display""
              ]
            },
            {
              ""type"": ""Point"",
              ""coordinates"": [
                42.589264732421,
                -74.3212407793175
              ],
              ""calculationMethod"": ""Rooftop"",
              ""usageTypes"": [
                ""Route""
              ]
            }
          ],
          ""matchCodes"": [
            ""Ambiguous""
          ]
        },
        {
          ""__type"": ""Location:http://schemas.microsoft.com/search/local/ws/rest/v1"",
          ""bbox"": [
            41.1190572824293,
            -73.4229365651287,
            41.1267827175707,
            -73.4092634348713
          ],
          ""name"": ""123 W Main St, Norwalk, CT 06851"",
          ""point"": {
            ""type"": ""Point"",
            ""coordinates"": [
              41.12292,
              -73.4161
            ]
          },
          ""address"": {
            ""addressLine"": ""123 W Main St"",
            ""adminDistrict"": ""CT"",
            ""adminDistrict2"": ""Fairfield"",
            ""countryRegion"": ""United States"",
            ""formattedAddress"": ""123 W Main St, Norwalk, CT 06851"",
            ""locality"": ""Norwalk"",
            ""postalCode"": ""06851""
          },
          ""confidence"": ""Medium"",
          ""entityType"": ""Address"",
          ""geocodePoints"": [
            {
              ""type"": ""Point"",
              ""coordinates"": [
                41.12292,
                -73.4161
              ],
              ""calculationMethod"": ""Rooftop"",
              ""usageTypes"": [
                ""Display""
              ]
            },
            {
              ""type"": ""Point"",
              ""coordinates"": [
                41.1231211183589,
                -73.4160720856487
              ],
              ""calculationMethod"": ""Rooftop"",
              ""usageTypes"": [
                ""Route""
              ]
            }
          ],
          ""matchCodes"": [
            ""Ambiguous""
          ]
        },
        {
          ""__type"": ""Location:http://schemas.microsoft.com/search/local/ws/rest/v1"",
          ""bbox"": [
            41.1190830885324,
            -73.4225082566583,
            41.1268085236738,
            -73.408835121025
          ],
          ""name"": ""123 Main St, Norwalk, CT 06851"",
          ""point"": {
            ""type"": ""Point"",
            ""coordinates"": [
              41.1229458061031,
              -73.4156716888417
            ]
          },
          ""address"": {
            ""addressLine"": ""123 Main St"",
            ""adminDistrict"": ""CT"",
            ""adminDistrict2"": ""Fairfield"",
            ""countryRegion"": ""United States"",
            ""formattedAddress"": ""123 Main St, Norwalk, CT 06851"",
            ""locality"": ""Norwalk"",
            ""postalCode"": ""06851""
          },
          ""confidence"": ""Medium"",
          ""entityType"": ""Address"",
          ""geocodePoints"": [
            {
              ""type"": ""Point"",
              ""coordinates"": [
                41.1229458061031,
                -73.4156716888417
              ],
              ""calculationMethod"": ""InterpolationOffset"",
              ""usageTypes"": [
                ""Display""
              ]
            },
            {
              ""type"": ""Point"",
              ""coordinates"": [
                41.1229616555711,
                -73.4156158277734
              ],
              ""calculationMethod"": ""Interpolation"",
              ""usageTypes"": [
                ""Route""
              ]
            }
          ],
          ""matchCodes"": [
            ""Ambiguous""
          ]
        }
      ]
    }
  ],
  ""statusCode"": 200,
  ""statusDescription"": ""OK"",
  ""traceId"": ""95ef4c8077004fd59cda27be0effea4a|BN20260231|7.7.0.0|""
}"
        Dim strip = New Strip()
        Try
            response = strip.Run(response)
        Catch ex As Exception
            response = ex.Message
        End Try
        HttpContext.Current.Response.Write(response)
    End Sub

End Class