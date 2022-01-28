# FilesWithAspose
This is a sample POC making use of Aspose and OpenXML which demonstrates the corrupted file problem and resolution plan for it.
-------------------------------------------------------------------------------------------------------------

Objective: This is a simple console application to validate, manipulate the file using Aspose and OpenXML

-------------------------------------------------------------------------------------------------------------

Steps To Run The Console Application:
> To Start With, Set "CorruptFilesValidationWithAspose" project as the Start-up Project. 
> Open Package Manager Console and Run This Command to add dependencies, missing this step will cause Build Error.
	- Install-Package Aspose.PDF -Version 19.8.0
	- Install-Package Aspose.Words -Version 19.8.0
	- Install-Package DocumentFormat.OpenXml -Version 2.14.0
	- Install-Package System.IO.Compression -Version 4.3.0

-------------------------------------------------------------------------------------------------------------

Assets/ Results of Console App:
> ReadMe.md
> CorruptFilesValidationWithAspose\Payload\

-------------------------------------------------------------------------------------------------------------

Command to Install Aspose.Words Library with different versions
> Install-Package Aspose.Words -Version 21.10.0
> Install-Package Aspose.Words -Version 21.9.0
> Install-Package Aspose.Words -Version 20.1.0

-------------------------------------------------------------------------------------------------------------

Problem Statement:
- We accept user payload as Word and Pdf document and keep it in our db or store
- Before this, one round of concrete validation has to be performed wheather the document is corrupted or not
- If the document happens to be corrupted, then we should not allow that document to be stored in the system and update user to re-sumbit the non-corrupted document
- This is a sample POC making use of Aspose and OpenXML which demonstrates the above problem and resolution plan for it.

-------------------------------------------------------------------------------------------------------------
