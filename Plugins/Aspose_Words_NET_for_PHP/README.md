# Aspose-Words-NET-for-PHP
Project Aspose.Words .NET for PHP shows how different tasks can be performed using Aspose.Words .NET APIs in PHP. This project is aimed to provide useful examples for PHP Developers who want to utilise Aspose.Words for .NET in their PHP Projects.

### System Requirements
* IIS with PHP and PHP Manager installed.
* Aspose.Total APIs.

### Supported Platforms
* PHP 5.3 or above
* Windows OS

### How to configure the source code on Windows Platform
#### 1. Register the dll file e.g. Aspose.Words.dll.
Register the dll file e.g. Aspose.Words.dll.
C:\Windows\Microsoft.NET\Framework\v2.0.50727>regasm c:\words\Aspose.Words.dll /codebase
Microsoft (R) .NET Framework Assembly Registration Utility 2.0.50727.7905
Copyright (C) Microsoft Corporation 1998-2004. All rights reserved.
Types registered successfully

#### 2. Enable COM and DOTNET Extensions in PHP
In IIS open PHP Manager and then Click ‘Enable to Disable and Extension‘. Find php_com_dotnet.dll and make sure it is enabled.

#### 3. Configure Aspose.Cells Java for PHP Examples
* Clone the Repository
```
git clone git@github.com:asposemarketplace/Aspose_Words_NET_for_PHP.git
```
* Setup the project using composer
```
composer install
```

Read more about composer on http://www.getcomposer.org
