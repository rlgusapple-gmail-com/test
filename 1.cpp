size_t ExtractFileFromOOXML(std::tstring strFilePath, std::tstring strStreamName, std::vector<std::tstring>& vecExtractedFile)
{
	std::tstring strDirectory = GetCurrentDirectory();
	std::vector<std::tstring> vecFiles;
	DWORD dwRet = Unzip(strFilePath, strStreamName, strDirectory, vecFiles);
	if( EC_SUCCESS != dwRet )
		return 0U;

	size_t i;
	for(i=0; i<vecFiles.size(); i++)
	{
		std::vector<BYTE> vecFileData;
		ReadFileContents(vecFiles[i], vecFileData);

		if( vecFileData.size() > 4 )
		{
			// for compound
			const DWORD dwDocHeader = 0xE011CFD0;
			if( 0 == memcmp(&dwDocHeader, &vecFileData[0], sizeof(dwDocHeader)) )
			{
				vecExtractedFile.push_back(vecFiles[i]);
				ExtractFileFromCompound(vecFileData, 0, vecExtractedFile);
				continue;
			}

			// for ooxml
			char szOoxmlHeader[3] = {static_cast<char>(vecFileData[0]), static_cast<char>(vecFileData[1]), 0 };
			if (MakeUpper(szOoxmlHeader) == "PK")
			{
				vecExtractedFile.push_back(vecFiles[i]);
				ExtractFileFromOOXML(vecFiles[i], strStreamName, vecExtractedFile);
				continue;
			}

			// for HWPX
			DWORD dwFileSize = *(DWORD*)&vecFileData[0];
			if( dwFileSize + 4 <= vecFileData.size() )
			{
				if (0 == memcmp(&dwDocHeader, &vecFileData[4], sizeof(dwDocHeader)))
					ExtractFileFromCompound(vecFileData, 4, vecExtractedFile);
			}
		}

		core::DeleteFile(vecFiles[i].c_str());
	}
	
	return vecExtractedFile.size();
}

bool CheckOOXML(std::tstring strSample)
{
	std::set<std::tstring> setNodes;
	FileListFromZIP(strSample, setNodes);

	if (!setNodes.empty())
		return CheckNodes(setNodes);

	return CheckCompound(strSample);
}

void FileListFromZIP(std::tstring strZipFile, std::set<std::tstring>& vecFiles)
{
	ZRESULT nRet;
	HZIP hZip = OpenZip(strZipFile.c_str(), NULL);

	ZIPENTRY stSummary = { 0, };
	nRet = GetZipItem(hZip, -1, &stSummary);

	int nItemCount = stSummary.index;
	int i;
	for (i = 0; i < nItemCount; i++)
	{
		ZIPENTRY stEntry = { 0, };
		nRet = GetZipItem(hZip, i, &stEntry);
		if (ZR_OK != nRet)
			continue;

		vecFiles.insert(stEntry.name);
	}

	CloseZip(hZip);
}

bool CheckCompound(std::tstring strSample)
{
	std::set<std::tstring> setNodes = QueryComNodes(strSample);
	return CheckNodes(setNodes);
}

bool CheckNodes(const std::set<std::tstring>& setNodes)
{
	bool bCondition1 = setNodes.find(TEXT("/Root Entry/EncryptedPackage")) != setNodes.end();
	bool bCondition2 = setNodes.find(TEXT("/Root Entry/EncryptionInfo")) != setNodes.end();
	return bCondition1 && bCondition2;
}

bool IsEncryptedDoc(std::tstring strSample)
{
	ECODE nRet = EC_SUCCESS;

	try
	{
		ST_SLE_COMMAND_DISTINCT_RETURN stFileType;
		ECODE nRet = DistinctFileType(strSample.c_str(), stFileType);
		if (EC_SUCCESS != nRet)
			throw exception_format("DistinctFileType operation failure, %d", nRet);

		if (SLE_FILE_TYPE_DOC == stFileType.nFileCategory)
		{
			switch (stFileType.nFileEXTCode)
			{
			case SLE_DOC_DOC:	return CheckDOC(strSample);
			case SLE_DOC_HWP:	return CheckHWP(strSample);
			case SLE_DOC_PPT:	return CheckPPT(strSample);
			case SLE_DOC_XLS:	return CheckXLS(strSample);
			case SLE_DOC_DOCX:
			case SLE_DOC_HWPX:
			case SLE_DOC_PPTX:
			case SLE_DOC_XLSX:
				return CheckOOXML(strSample);
			}
		}

		return false;
	}
	catch (std::exception& e)
	{
		Log_Error("%s", e.what());
		return false;
	}
}



