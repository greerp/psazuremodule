create view euc_files as 

	/*
	  Query to extract scanned files and their attributes 
  	  from Converter Technology Database

	  Written by: Paul Greer, RedPixie UK Ltd 11th March 2016 

	*/
	with fattr_cte(FileId, AttributeName,AttributeValue) as
		(
		select v.FileId,m.AttributeName,v.AttributeValue 
		from FileAttributeValues v 
			inner join FileAttributesMaster m on m.AttributeID=v.AttributeID
		),
		fhas_cte(fileId, basefolder, HasVBA, HasCharts, HasPasswords, HasLinks, HasActiveX, Loc, BrokenRef, LateBinding, [Hash]) as
        (
		  select f.FileID as fileId,
  		    substring(f.FilePath,44, charindex('\',substring(f.FilePath,44,1000))-1) as BaseFolder,
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasVBA') as HasVBA, 
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasCharts') as HasCharts, 
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasPasswords') as HasPasswords, 
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasLinks') as HasLinks,
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasActiveX') as HasActiveX,
			vba.Loc as Loc,
			(case when exists(select 1 from xlsreferences xref where f.fileid=xref.fileid and xref.IsBroken=1) then 1 else 0 end) as BrokenRef,
			vba.createobject as LateBinding,
			vba.[Hash] as [Hash]
		  from Files f left outer join xlsvbaloc vba on f.FileID=vba.FileId 
        )
	    select 
           f1.*, 
           f2.basefolder, f2.HasActiveX, f2.hasVBA, f2.HasCharts, f2.HasLinks, f2.HasPasswords, f2.Loc, f2.BrokenRef,f2.LateBinding,
           ragstatus=
		   case when f2.HasActiveX!='No' or f2.HasVBA!='No' or f2.HasCharts!='No' then 'RED'
		        when f2.HasLinks='Yes' then 'AMBER'
		        else 'GREEN' end,
		   userragstatus=
		   case when f2.HasActiveX='Yes' and f2.brokenRef=1 then 'RED'
		        when f2.Loc > 0 then 'AMBER'
		        when f2.HasVba='Unknown'  then 'AMBER'
		        when f2.HasVBA='Yes' and f2.Loc IS NULL then 'AMBER'
		        when f2.HasCharts!='No'  then 'AMBER'
		        else 'GREEN' end, 	
		   complexity=
		   case when Loc > 1000 then 'HIGH'
		        when Loc > 100 and Loc < 1000 then 'MED'
		        when HasVBA<>'No' and Loc is null then 'UNKNOWN'
		        else 'LOW' end,
		   f2.[Hash]
        from Files f1, fhas_cte f2 where f1.FileID=f2.fileId

