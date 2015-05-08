Imports HtmlAgilityPack
Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Configuration.ConfigurationManager
Imports Microsoft.WindowsAzure
Imports System.Threading
Imports System.Linq

Module Module1
    Dim VisitedLinks As New List(Of String)
    Dim MaxDepth As Integer

    Sub Main()
        'Get the max link depth to traverse from App Settings
        MaxDepth = System.Configuration.ConfigurationManager.AppSettings("ProjectNamiCacheLoader.MaxDepth")

        'Locate all URLs in the Site List, we could be dealing with a multisite
        For Each Site As String In System.Configuration.ConfigurationManager.AppSettings("ProjectNamiCacheLoader.SiteList").Split(",")
            'Prepend a scheme if only a host name was provided
            If Not Site.StartsWith("http://") And Not Site.StartsWith("https://") Then
                Site = "http://" & Site
            End If
            VisitLink(Site, 0)
        Next
    End Sub

    Sub VisitLink(LinkURL As String, LinkDepth As Integer)
        Try
            'Create URI object from string
            Dim ThisURI As New Uri(LinkURL)

            'Get the page as a simple WebResponse so client-side tracking elements (pixels, JavaScript, etc) will not be triggered
            Dim ThisResponse As WebResponse = GetResponse(LinkURL)

            'Check for a clean response
            If CType(ThisResponse, HttpWebResponse).StatusCode = 200 Then
                'Add the link to our history
                VisitedLinks.Add(LinkURL)
                Console.Out.WriteLine("Visited - " & LinkURL)

                'Get the HTML payload
                Dim PageHTML As String = New StreamReader(ThisResponse.GetResponseStream).ReadToEnd

                'If we haven't reached our Max Depth...
                If LinkDepth < MaxDepth Then
                    Try
                        'Create HtmlDocument from the payload
                        Dim ThisPage As HtmlDocument = ParsePage(PageHTML)

                        If Not IsNothing(ThisPage) Then
                            'Extract links from A tags with valid HREF values
                            Dim PageLinks = From Links As HtmlNode In ThisPage.DocumentNode.SelectNodes("//a[@href]")
                                            Where (Links.Name = "a" And Not IsDBNull(Links.Attributes("href")) And Not Links.Attributes("href").Value.StartsWith("#"))
                                            Select New With {.url = Links.Attributes("href").Value}

                            Dim CleanLinks As New List(Of String)

                            For Each Link In PageLinks
                                'If the link is relative, prepend the scheme and site name
                                If Link.url.StartsWith("/") Then
                                    Link.url = ThisURI.Scheme & "://" & ThisURI.DnsSafeHost & Link.url
                                End If
                                Dim LinkURI As New Uri(Link.url)
                                'Only prepare to follow the link if it points to the same site
                                If LinkURI.DnsSafeHost = ThisURI.DnsSafeHost Then
                                    CleanLinks.Add(LinkURI.GetLeftPart(UriPartial.Query))
                                End If
                            Next

                            'Dedupe the links, no need to visit the same page more than once
                            Dim DistinctLinks As List(Of String) = CleanLinks.Distinct.ToList

                            For Each ThisLink As String In DistinctLinks
                                Try
                                    'Only attempt to follow the link if we haven't visited it yet
                                    If Not VisitedLinks.Any(Function(str) ThisLink = str) Then
                                        'Don't follow links into wp-admin
                                        If ThisLink.ToLower.Contains("/wp-admin/") Then
                                            VisitedLinks.Add(ThisLink)
                                        Else
                                            VisitLink(ThisLink, LinkDepth + 1)
                                        End If
                                    End If
                                Catch ex As Exception
                                    Console.Out.WriteLine("ERROR Visit a link - " & ex.Message & ex.StackTrace)
                                End Try
                            Next
                        End If
                    Catch ex As Exception
                        Console.Out.WriteLine("ERROR Parse Links - " & ex.Message & ex.StackTrace)
                        Thread.Sleep(20000)
                    End Try

                End If
            Else
                Console.Out.WriteLine("Attempted - " & LinkURL & " - Code " & CType(ThisResponse, HttpWebResponse).StatusCode)
            End If
        Catch ex As Exception
            'Total faliure, add the link to our history so we don't attempt it again
            VisitedLinks.Add(LinkURL)
            Console.Out.WriteLine("ERROR Unable to visit " & LinkURL & " - " & ex.Message & ex.StackTrace)
        End Try
    End Sub

    'Dedicated function for performing WebRequests so each request will fall out of scope and release resources easier
    Function GetResponse(LinkURL As String) As WebResponse
        Try
            Dim ThisRequest As WebRequest = WebRequest.Create(LinkURL)
            'Add the BypassKey found in App Settings to the User Agent, so we can notify the Blob Cache Front End to always let us through
            CType(ThisRequest, HttpWebRequest).UserAgent &= Space(1) & System.Configuration.ConfigurationManager.AppSettings("ProjectNamiBlobCache.BypassKey")
            Return ThisRequest.GetResponse
        Catch ex As Exception
            Console.Out.WriteLine("ERROR Get Response - " & ex.Message & ex.StackTrace)
            Return Nothing
        End Try
    End Function

    Function ParsePage(HTMLString As String) As HtmlDocument
        Try
            Dim ThisPage As New HtmlDocument
            ThisPage.LoadHtml(HTMLString)
            Return ThisPage
        Catch ex As Exception
            Console.Out.WriteLine("ERROR Parse Page - " & ex.Message & ex.StackTrace)
            Return Nothing
        End Try
    End Function


End Module
