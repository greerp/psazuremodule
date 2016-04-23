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
		fhas_cte(fileId, basefolder, HasVBA, HasCharts, HasPasswords, HasLinks, HasActiveX) as
        (
		  select f.FileID as fileId,
  		    substring(f.FilePath,44, charindex('\',substring(f.FilePath,44,1000))-1) as BaseFolder,
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasVBA') as HasVBA, 
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasCharts') as HasCharts, 
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasPasswords') as HasPasswords, 
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasLinks') as HasLinks,
		    (select AttributeValue from fattr_cte f1 where f.FileID=f1.FileId and f1.AttributeName='HasActiveX') as HasActiveX,
     		    (select loc from xlsvbaloc loc where loc.fileid = f.FileID) as Loc
		  from Files f  
        )
	    select 
           f1.*, 
           f2.basefolder, f2.HasActiveX, f2.hasVBA, f2.HasCharts, f2.HasLinks, f2.HasPasswords, f2.Loc
           ragstatus=case when f2.HasActiveX!='No' or f2.HasVBA!='No' or f2.HasCharts!='No' then 'RED'
		        when f2.HasLinks='Yes' then 'AMBER'
		        else 'GREEN' end,
		   userragstatus=case when f2.HasActiveX='Yes' then 'RED'
		                      when f2.HasVBA!='No' or f2.HasCharts!='No'  then 'AMBER'
		        else 'GREEN' end 				
        from Files f1, fhas_cte f2 where f1.FileID=f2.fileId