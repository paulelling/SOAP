<%@ WebService Language="C#" Class="HelloWorld" %>

using System;
using System.Web.Services;

public class HelloWorld : WebService
{
  [WebMethod] public String SayHelloWorld()
  {
    return "Hello World";
  }
}