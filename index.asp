<!--#include file="class_aws_s3.asp"-->
<%



	Set amz = New amazon
	'upload file'
	'###################'
	amz.s3_LocalFile				= "c:\path\file.zip"
	amz.s3_RemoteFile				= "path/file.zip"
	amz.s3_Bucket					= "bucketname"
	result = amz.s3_Upload()

	amz.s3_DeleteLocalFile()
	
	response.write result
	'###################'
	Set amz = Nothing





	Set amz = New amazon
	'upload stream file'
	'###################'
	amz.s3_LocalFile				= "c:\path\file.zip"
	amz.s3_RemoteFile				= "path/file.zip"
	amz.s3_Bucket					= "bucketname"
	result = amz.s3_Upload()
	response.write result
	'###################'
	Set amz = Nothing





	Set amz = New amazon
	'upload stream file'
	'###################'
	amz.s3_UploadBinary				= binary
	amz.s3_RemoteFile				= "path/test.zip"
	amz.s3_Bucket					= "bucketname"
	result = amz.s3_UploadBinary()
	response.write result
	'###################'
	Set amz = Nothing




	Set amz = New amazon
	'download file'
	'###################'
	amz.s3_LocalFile				= "c:\path\test.pdf"
	amz.s3_RemoteFile				= "test.pdf"
	amz.s3_Bucket					= "bucketname"
	result = amz.s3_Download()
	response.write result
	'###################'
	Set amz = Nothing



	Set amz = New amazon
	'stream file'
	'###################'
	amz.s3_RemoteFile				= "test.pdf"
	amz.s3_Bucket					= "bucketname"
	amz.s3_OutFileName				= "test.pdf"
	amz.s3_StreamToBrowser()
	'###################'
	Set amz = Nothing


	Set amz = New amazon
	'delete local file'
	'###################'
	amz.s3_LocalFile				= "test.pdf"
	amz.s3_DeleteLocalFile()
	'###################'
	Set amz = Nothing



	Set amz = New amazon
	'delete remote file'
	'###################'
	amz.s3_RemoteFile				= "test.pdf"
	amz.s3_Bucket					= "bucketname"
	amz.s3_Delete()
	'###################'
	Set amz = Nothing




	





%>