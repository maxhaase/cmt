IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'Smart')
	DROP DATABASE [Smart]
GO

CREATE DATABASE [Smart]  ON (NAME = N'SMART_Data', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL\data\SMART_Data.MDF' , SIZE = 2, FILEGROWTH = 10%) LOG ON (NAME = N'SMART_Log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL\data\SMART_Log.LDF' , SIZE = 2, FILEGROWTH = 10%)
 COLLATE SQL_Latin1_General_CP1_CI_AS
GO

exec sp_dboption N'Smart', N'autoclose', N'false'
GO

exec sp_dboption N'Smart', N'bulkcopy', N'false'
GO

exec sp_dboption N'Smart', N'trunc. log', N'false'
GO

exec sp_dboption N'Smart', N'torn page detection', N'true'
GO

exec sp_dboption N'Smart', N'read only', N'false'
GO

exec sp_dboption N'Smart', N'dbo use', N'false'
GO

exec sp_dboption N'Smart', N'single', N'false'
GO

exec sp_dboption N'Smart', N'autoshrink', N'false'
GO

exec sp_dboption N'Smart', N'ANSI null default', N'false'
GO

exec sp_dboption N'Smart', N'recursive triggers', N'false'
GO

exec sp_dboption N'Smart', N'ANSI nulls', N'false'
GO

exec sp_dboption N'Smart', N'concat null yields null', N'false'
GO

exec sp_dboption N'Smart', N'cursor close on commit', N'false'
GO

exec sp_dboption N'Smart', N'default to local cursor', N'false'
GO

exec sp_dboption N'Smart', N'quoted identifier', N'false'
GO

exec sp_dboption N'Smart', N'ANSI warnings', N'false'
GO

exec sp_dboption N'Smart', N'auto create statistics', N'true'
GO

exec sp_dboption N'Smart', N'auto update statistics', N'true'
GO

use [Smart]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CustomersCases_Cases]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CustomersCases] DROP CONSTRAINT FK_CustomersCases_Cases
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CategoriesImages_Categories]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CategoriesImages] DROP CONSTRAINT FK_CategoriesImages_Categories
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HousesCategories_Categories]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[HousesCategories] DROP CONSTRAINT FK_HousesCategories_Categories
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FilesPhrases_Files]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FilesPhrases] DROP CONSTRAINT FK_FilesPhrases_Files
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HousesCategories_Houses]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[HousesCategories] DROP CONSTRAINT FK_HousesCategories_Houses
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HousesImages_Houses]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[HousesImages] DROP CONSTRAINT FK_HousesImages_Houses
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HousesLocations_Houses]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[HousesLocations] DROP CONSTRAINT FK_HousesLocations_Houses
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HousesTypes_Houses]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[HousesTypes] DROP CONSTRAINT FK_HousesTypes_Houses
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CategoriesImages_Images]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CategoriesImages] DROP CONSTRAINT FK_CategoriesImages_Images
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HousesImages_Images]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[HousesImages] DROP CONSTRAINT FK_HousesImages_Images
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_LocationsImages_Images]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[LocationsImages] DROP CONSTRAINT FK_LocationsImages_Images
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ThumbnailsImages_Images]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PhotosImages] DROP CONSTRAINT FK_ThumbnailsImages_Images
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Customers_Languages]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Customers] DROP CONSTRAINT FK_Customers_Languages
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Text_Languages]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Text] DROP CONSTRAINT FK_Text_Languages
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HousesLocations_Locations]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[HousesLocations] DROP CONSTRAINT FK_HousesLocations_Locations
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_LocationsImages_Locations]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[LocationsImages] DROP CONSTRAINT FK_LocationsImages_Locations
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_ThumbnailsImages_Thumbnails]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PhotosImages] DROP CONSTRAINT FK_ThumbnailsImages_Thumbnails
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_FilesPhrases_Phrases]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[FilesPhrases] DROP CONSTRAINT FK_FilesPhrases_Phrases
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PhrasesTypes_Phrases]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PhrasesTypes] DROP CONSTRAINT FK_PhrasesTypes_Phrases
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Text_Phrases]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Text] DROP CONSTRAINT FK_Text_Phrases
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HousesTypes_Types]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[HousesTypes] DROP CONSTRAINT FK_HousesTypes_Types
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_PhrasesTypes_Types]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[PhrasesTypes] DROP CONSTRAINT FK_PhrasesTypes_Types
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_CustomersCases_Customers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[CustomersCases] DROP CONSTRAINT FK_CustomersCases_Customers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cases_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Cases_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cases_Fetch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Cases_Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CustomersCases_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[CustomersCases_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CustomersCases_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[CustomersCases_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CategoriesImages_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[CategoriesImages_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CategoriesImages_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[CategoriesImages_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Customers_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Customers_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Customers_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Customers_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Customers_Update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Customers_Update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FilesPhrases_Fetch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[FilesPhrases_Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FilesPhrases_add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[FilesPhrases_add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FilesPhrases_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[FilesPhrases_delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FilesPhrases_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[FilesPhrases_update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Files_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Files_delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HousesImages_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[HousesImages_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HousesImages_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[HousesImages_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Images_Fetch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Images_Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LocationsImages_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LocationsImages_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LocationsImages_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LocationsImages_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PhotosImages_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PhotosImages_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PhotosImages_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PhotosImages_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PhrasesTypes_add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PhrasesTypes_add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PhrasesTypes_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PhrasesTypes_delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PhrasesTypes_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PhrasesTypes_update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Phrases_Fetch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Phrases_Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Phrases_add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Phrases_add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Phrases_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Phrases_delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Phrases_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Phrases_update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Text_Fetch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Text_Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Text_add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Text_add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Text_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Text_delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Text_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Text_update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cases_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Cases_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cases_Update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Cases_Update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Files_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Files_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Files_Fetch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Files_Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Files_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Files_update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Houses_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Houses_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Houses_Fetch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Houses_Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Houses_Update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Houses_Update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Images_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Images_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Images_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Images_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Images_Update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Images_Update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Languages_Fetch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Languages_Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Languages_add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Languages_add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Languages_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Languages_delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Languages_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Languages_update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Locations_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Locations_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Locations_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Locations_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Locations_Delete_]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Locations_Delete_]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Locations_Update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Locations_Update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Locations_Update_]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Locations_Update_]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Photos_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Photos_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Photos_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Photos_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Photos_Update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Photos_Update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Types_ delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Types_ delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Types_add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Types_add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Types_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Types_update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Users_add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Users_add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Users_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Users_delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Users_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Users_update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Categories_Add]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Categories_Add]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Categories_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Categories_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Categories_Delete_]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Categories_Delete_]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Categories_Update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Categories_Update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Categories_Update_]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Categories_Update_]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fetch]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[Fetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ImagesFetch]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[ImagesFetch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CustomersCases]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CustomersCases]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CategoriesImages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CategoriesImages]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Customers]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Customers]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FilesPhrases]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FilesPhrases]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HousesCategories]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HousesCategories]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HousesImages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HousesImages]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HousesLocations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HousesLocations]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HousesTypes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HousesTypes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LocationsImages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LocationsImages]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PhotosImages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PhotosImages]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PhrasesTypes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PhrasesTypes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Text]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Text]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cases]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cases]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Categories]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Categories]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Distances]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Distances]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Files]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Files]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Houses]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Houses]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Images]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Images]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Languages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Languages]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Locations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Locations]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Photos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Photos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Phrases]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Phrases]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Types]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Types]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Users]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Users]
GO

CREATE TABLE [dbo].[Cases] (
	[lngCaseId] [int] IDENTITY (1, 1) NOT NULL ,
	[strCase] [varchar] (8000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Categories] (
	[lngCategoryId] [int] IDENTITY (1, 1) NOT NULL ,
	[strCategory] [varchar] (50) COLLATE Latin1_General_CS_AS NOT NULL ,
	[lngPriceFrom] [int] NULL ,
	[lngPriceTo] [int] NULL ,
	[strNotes] [varchar] (8000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Distances] (
	[lngDistance] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Files] (
	[lngFileId] [int] IDENTITY (1, 1) NOT NULL ,
	[strFileName] [varchar] (1600) COLLATE Latin1_General_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Houses] (
	[lngHouseId] [int] IDENTITY (1, 1) NOT NULL ,
	[lngCategoryId] [int] NOT NULL ,
	[lngLocationId] [int] NOT NULL ,
	[lngTypeId] [int] NOT NULL ,
	[strTerms] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[lngPrice] [int] NULL ,
	[lngArea] [int] NULL ,
	[lngBedrooms] [int] NULL ,
	[lngBathrooms] [int] NULL ,
	[blGarage] [bit] NULL ,
	[blRoofTerrace] [bit] NULL ,
	[blSwimmingPool] [bit] NULL ,
	[lngPatioArea] [int] NULL ,
	[lngDistanceBeach] [int] NULL ,
	[lngDistanceGolf] [int] NULL ,
	[lngDistanceAirport] [int] NULL ,
	[lngDistanceCentrum] [int] NULL ,
	[lngYearBuilt] [int] NULL ,
	[strNotes] [varchar] (8000) COLLATE Latin1_General_CS_AS NULL ,
	[strAddress] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[strCity] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[strProvince] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[strCountry] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[strOwnerName] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[strOwnerLastName] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[strOwnerPhone] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[strOwnerFax] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[strOwnerEmail] [varchar] (50) COLLATE Latin1_General_CS_AS NULL ,
	[blFurniture] [bit] NULL ,
	[lngFloorCount] [int] NULL ,
	[lngFloor] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Images] (
	[lngImageId] [int] IDENTITY (1, 1) NOT NULL ,
	[strImage] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Languages] (
	[lngLanguageId] [int] NOT NULL ,
	[lngPhraseId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Locations] (
	[lngLocationId] [int] IDENTITY (1, 1) NOT NULL ,
	[strLocation] [varchar] (50) COLLATE Latin1_General_CS_AS NOT NULL ,
	[strNotes] [varchar] (8000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Photos] (
	[lngPhotoId] [int] IDENTITY (1, 1) NOT NULL ,
	[strPhoto] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Phrases] (
	[lngPhraseId] [int] NOT NULL ,
	[strPhrase] [varchar] (8000) COLLATE Latin1_General_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Types] (
	[lngTypeId] [int] IDENTITY (1, 1) NOT NULL ,
	[strType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Users] (
	[lngUserId] [int] IDENTITY (1, 1) NOT NULL ,
	[strUserName] [varchar] (50) COLLATE Latin1_General_CS_AS NOT NULL ,
	[strFirstName] [varchar] (50) COLLATE Latin1_General_CS_AS NOT NULL ,
	[strLastName] [varchar] (50) COLLATE Latin1_General_CS_AS NOT NULL ,
	[strEmail] [varchar] (50) COLLATE Latin1_General_CS_AS NOT NULL ,
	[strPassword] [varchar] (50) COLLATE Latin1_General_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CategoriesImages] (
	[lngCategoryId] [int] NULL ,
	[lngImageId] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Customers] (
	[lngCustomerId] [int] IDENTITY (1, 1) NOT NULL ,
	[strFirstName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[strLastName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[strAddress] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strCity] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strState] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strZip] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strCountry] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strFax] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strEmail] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strOrganizationNr] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strWebSite] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[strNotes] [varchar] (8000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[lngLanguageId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FilesPhrases] (
	[lngFileId] [int] NOT NULL ,
	[lngPhraseId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HousesCategories] (
	[lngHouseId] [int] NOT NULL ,
	[lngCategoryId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HousesImages] (
	[lngImageId] [int] NOT NULL ,
	[lngHouseId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HousesLocations] (
	[lngHouseId] [int] NOT NULL ,
	[lngLocationId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HousesTypes] (
	[lngHouseId] [int] NULL ,
	[lngTypeId] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LocationsImages] (
	[lngLocationId] [int] NOT NULL ,
	[lngImageId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PhotosImages] (
	[lngPhotoId] [int] NOT NULL ,
	[lngImageId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PhrasesTypes] (
	[lngPhraseTypeId] [int] NOT NULL ,
	[lngPhraseId] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Text] (
	[lngPhraseId] [int] NOT NULL ,
	[lngLanguageId] [int] NOT NULL ,
	[strPhrase] [varchar] (8000) COLLATE Latin1_General_CS_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CustomersCases] (
	[lngCustomerId] [int] NOT NULL ,
	[lngCaseId] [int] NOT NULL 
) ON [PRIMARY]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.[Fetch]
AS
SELECT     dbo.Houses.*
FROM         dbo.Houses INNER JOIN
                      dbo.HousesCategories ON dbo.Houses.lngHouseId = dbo.HousesCategories.lngHouseId INNER JOIN
                      dbo.HousesLocations ON dbo.Houses.lngHouseId = dbo.HousesLocations.lngHouseId INNER JOIN
                      dbo.HousesTypes ON dbo.Houses.lngHouseId = dbo.HousesTypes.lngHouseId INNER JOIN
                      dbo.Categories ON dbo.HousesCategories.lngCategoryId = dbo.Categories.lngCategoryId

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.ImagesFetch
AS
SELECT     dbo.Photos.strPhoto, dbo.Photos.lngPhotoId, dbo.PhotosImages.lngImageId, dbo.Images.strImage
FROM         dbo.PhotosImages INNER JOIN
                      dbo.Images ON dbo.PhotosImages.lngImageId = dbo.Images.lngImageId INNER JOIN
                      dbo.LocationsImages ON dbo.Images.lngImageId = dbo.LocationsImages.lngImageId INNER JOIN
                      dbo.HousesImages ON dbo.Images.lngImageId = dbo.HousesImages.lngImageId INNER JOIN
                      dbo.Photos ON dbo.PhotosImages.lngPhotoId = dbo.Photos.lngPhotoId INNER JOIN
                      dbo.CategoriesImages ON dbo.Images.lngImageId = dbo.CategoriesImages.lngImageId
WHERE     (dbo.LocationsImages.lngLocationId = 1)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Categories_Add]
	(@lngCategoryId 	[int] OUTPUT,
	 @strCategory 	[varchar](50),
	 @lngPriceFrom 	[int],
	 @lngPriceTo 	[int],
	 @strNotes 	[int])

AS

 INSERT INTO  [Categories] 
	 ( [strCategory],
	 [lngPriceFrom],
	 [lngPriceTo],
	 [strNotes]) 
 
VALUES 
	( @strCategory,
	 @lngPriceFrom,
	 @lngPriceTo,
	 @strNotes)

SET @lngCategoryId = @@identity
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Categories_Delete]
	(@lngCategoryId 	[int])

AS DELETE [Categories] 

WHERE 
	( [lngCategoryId]	 = @lngCategoryId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Categories_Delete_]
	(@lngCategoryId 	[int])

AS DELETE [Categories] 

WHERE 
	( [lngCategoryId]	 = @lngCategoryId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Categories_Update]
	(@lngCategoryId 	[int],
	 @strCategory 	[varchar](50),
	 @lngPriceFrom 	[int],
	 @lngPriceTo 	[int])

AS UPDATE  [Categories] 

SET  [strCategory]	 = @strCategory,
	 [lngPriceFrom]	 = @lngPriceFrom,
	 [lngPriceTo]	 = @lngPriceTo 

WHERE 
	( [lngCategoryId]	 = @lngCategoryId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Categories_Update_]
	(@lngCategoryId 	[int],
	 @strCategory 	[varchar](50),
	 @lngPriceFrom 	[int],
	 @lngPriceTo 	[int],
	 @strNotes 	[int])

AS UPDATE [Categories] 

SET  [strCategory]	 = @strCategory,
	 [lngPriceFrom]	 = @lngPriceFrom,
	 [lngPriceTo]	 = @lngPriceTo,
	 [strNotes]	 = @strNotes 

WHERE 
	( [lngCategoryId]	 = @lngCategoryId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Cases_Add]
	(@strCase 	[varchar](8000),
	@lngCaseId [int] OUTPUT)

AS INSERT INTO [Cases] 
	 ( [strCase]) 
 
VALUES 
	( @strCase)



SET @lngCaseId = @@identity
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Cases_Update]
	(@lngCaseId 	[int],
	 @strCase 	[varchar](8000))

AS UPDATE [Cases] 

SET  [strCase]	 = @strCase 

WHERE 
	( [lngCaseId]	 = @lngCaseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Files_Add]
	(@strFileName 	[varchar](1000),
	@lngFileId 	[int]=0 output,
	@COUNT	[int]=0)

AS

select @COUNT = count(lngFileId) from Files where strFileName = @strFileName

if @COUNT = 0 

BEGIN
 	INSERT INTO [Files] 
	 	( [strFileName]) 
 	VALUES 
		( @strFileName)
END


IF @@ERROR = 0 
	BEGIN
		set @lngFileId = @@identity
	END
ELSE

		SET @lngFileId = 0
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



CREATE PROCEDURE [Files_Fetch]
	(@lngFileId [int])

AS

if @lngFileId <> 0 
	begin
	select * from [Files]  where lngFileId = @lngFileId
	end
else
	begin
	select * from [Files] 
	end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Files_update]
	(@lngFileId 	[int],
	 @strFileName 	[varchar](1000),
	@lngCount	[int]= 0 output)

AS 

SELECT @lngCount = COUNT(lngFileId) FROM Files WHERE lngFileId = @lngFileId

IF @lngCount <>  0
	BEGIN
		UPDATE [Files] 

		SET  [strFileName]	 = @strFileName 

		WHERE 
			( [lngFileId]	 = @lngFileId)
	END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Houses_Add]
	(@lngCategoryId 	[int],
	 @lngLocationId 	[int],
	 @lngTypeId 	[int],
	 @strTerms 	[varchar](50),
	 @lngPrice 	[int],
	 @lngArea 	[int],
	 @lngBedrooms 	[int],
	 @lngBathrooms 	[int],
	 @blGarage 	[bit],
	 @blRoofTerrace 	[bit],
	 @blSwimmingPool 	[bit],
	 @lngPatioArea 	[int],
	 @lngDistanceBeach 	[int],
	 @lngDistanceGolf 	[int],
	 @lngDistanceAirport 	[int],
	 @lngDistanceCentrum 	[int],
	 @lngYearBuilt 	[int],
	 @strNotes 	[varchar](8000),
	 @strAddress 	[varchar](50),
	 @strCity 	[varchar](50),
	 @strProvince 	[varchar](50),
	 @strCountry 	[varchar](50),
	 @strOwnerName 	[varchar](50),
	 @strOwnerLastName 	[varchar](50),
	 @strOwnerPhone 	[varchar](50),
	 @strOwnerFax 	[varchar](50),
	 @strOwnerEmail 	[varchar](50),
	 @blFurniture	[bit])

AS INSERT INTO [Houses] 
	 ( [lngCategoryId],
	 [lngLocationId],
	 [lngTypeId],
	 [strTerms],
	 [lngPrice],
	 [lngArea],
	 [lngBedrooms],
	 [lngBathrooms],
	 [blGarage],
	 [blRoofTerrace],
	 [blSwimmingPool],
	 [lngPatioArea],
	 [lngDistanceBeach],
	 [lngDistanceGolf],
	 [lngDistanceAirport],
	 [lngDistanceCentrum],
	 [lngYearBuilt],
	 [strNotes],
	 [strAddress],
	 [strCity],
	 [strProvince],
	 [strCountry],
	 [strOwnerName],
	 [strOwnerLastName],
	 [strOwnerPhone],
	 [strOwnerFax],
	 [strOwnerEmail],
	 [blFurniture]) 
 
VALUES 
	( @lngCategoryId,
	 @lngLocationId,
	 @lngTypeId,
	 @strTerms,
	 @lngPrice,
	 @lngArea,
	 @lngBedrooms,
	 @lngBathrooms,
	 @blGarage,
	 @blRoofTerrace,
	 @blSwimmingPool,
	 @lngPatioArea,
	 @lngDistanceBeach,
	 @lngDistanceGolf,
	 @lngDistanceAirport,
	 @lngDistanceCentrum,
	 @lngYearBuilt,
	 @strNotes,
	 @strAddress,
	 @strCity,
	 @strProvince,
	 @strCountry,
	 @strOwnerName,
	 @strOwnerLastName,
	 @strOwnerPhone,
	 @strOwnerFax,
	 @strOwnerEmail,
	 @blFurniture)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Houses_Fetch] 
@lngCategoryId [int], 
@lngLocationId [int],
@lngTypeId [int],
@strTerms [varchar](50),
@lngBedrooms [int],
@lngPrice [int]

AS

if @lngCategoryId <> -1
	if @lngLocationId <> -1
		if @lngTypeId <> -1
			if @strTerms <> "-1"
				if @lngBedrooms <> -1
					if @lngPrice <> -1
						BEGIN
						SELECT * FROM Houses WHERE lngCategoryId=@lngCategoryId AND lngLocationId=@lngLocationId AND lngTypeId=@lngTypeId AND strTerms=@strTerms AND lngBedrooms=@lngBedrooms AND lngPrice=@lngPrice
						END
				ELSE
				
						

SELECT * FROM HOUSES
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Houses_Update]
	(@lngHouseId 	[int],
	 @lngCategoryId 	[int],
	 @lngLocationId 	[int],
	 @lngTypeId 	[int],
	 @strTerms 	[varchar](50),
	 @lngPrice 	[int],
	 @lngArea	[int],
	 @lngBedrooms 	[int],
	 @lngBathrooms 	[int],
	 @blGarage 	[bit],
	 @blRoofTerrace 	[bit],
	 @blSwimmingPool 	[bit],
	 @lngPatioArea 	[int],
	 @lngDistanceBeach 	[int],
	 @lngDistanceGolf 	[int],
	 @lngDistanceAirport 	[int],
	 @lngDistanceCentrum 	[int],
	 @lngYearBuilt 	[int],
	 @strNotes 	[varchar](8000),
	 @strAddress 	[varchar](50),
	 @strCity 	[varchar](50),
	 @strProvince 	[varchar](50),
	 @strCountry 	[varchar](50),
	 @strOwnerName 	[varchar](50),
	 @strOwnerLastName 	[varchar](50),
	 @strOwnerPhone 	[varchar](50),
	 @strOwnerFax 	[varchar](50),
	 @strOwnerEmail 	[varchar](50),
	 @blFurniture	[bit])

AS UPDATE [SMART].[dbo].[Houses] 

SET  [lngCategoryId]	 = @lngCategoryId,
	 [lngLocationId]	 = @lngLocationId,
	 [lngTypeId]	 = @lngTypeId,
	 [strTerms]	 = @strTerms,
	 [lngPrice]	 = @lngPrice,
	 [lngArea]	 = @lngArea,
	 [lngBedrooms]	 = @lngBedrooms,
	 [lngBathrooms]	 = @lngBathrooms,
	 [blGarage]	 = @blGarage,
	 [blRoofTerrace]	 = @blRoofTerrace,
	 [blSwimmingPool]	 = @blSwimmingPool,
	 [lngPatioArea]	 = @lngPatioArea,
	 [lngDistanceBeach]	 = @lngDistanceBeach,
	 [lngDistanceGolf]	 = @lngDistanceGolf,
	 [lngDistanceAirport]	 = @lngDistanceAirport,
	 [lngDistanceCentrum]	 = @lngDistanceCentrum,
	 [lngYearBuilt]	 = @lngYearBuilt,
	 [strNotes]	 = @strNotes,
	 [strAddress]	 = @strAddress,
	 [strCity]	 = @strCity,
	 [strProvince]	 = @strProvince,
	 [strCountry]	 = @strCountry,
	 [strOwnerName]	 = @strOwnerName,
	 [strOwnerLastName]	 = @strOwnerLastName,
	 [strOwnerPhone]	 = @strOwnerPhone,
	 [strOwnerFax]	 = @strOwnerFax,
	 [strOwnerEmail]	 = @strOwnerEmail,
	 [blFurniture]  = @blFurniture

WHERE 
	( [lngHouseId]	 = @lngHouseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Images_Add]
	(@lngImageId 	[int] output,
	 @strImage 	[varchar](2000))

AS INSERT INTO [Images] 
	 ([strImage]) 
 
VALUES 
	(  @strImage)

set @lngImageId = @@identity
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Images_Delete]
	(@lngImageId 	[int])

AS DELETE [Images] 

WHERE 
	( [lngImageId]	 = @lngImageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Images_Update]
	(@lngImageId 	[int],
	 @strImage 	[varchar](2000))

AS UPDATE [Images] 

SET  [strImage]	 = @strImage 

WHERE 
	( [lngImageId]	 = @lngImageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Languages_Fetch]  AS


BEGIN
	SELECT * FROM Languages
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Languages_add]
	( @lngPhraseId 	[int],
	@lngLanguageId 	[int] output)

AS
BEGIN
SELECT @lngLanguageId= MAX(lngLanguageId)+1 FROM Languages
END

BEGIN
 INSERT INTO [Languages]  ( [lngLanguageId], [lngPhraseId])  VALUES ( @lngLanguageId,	 @lngPhraseId)
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Languages_delete]
	(@lngLanguageId 	[int])

AS DELETE [Languages] 

WHERE 
	( [lngLanguageId]	 = @lngLanguageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Languages_update]
	(@lngLanguageId 	[int],
	 @lngPhraseId 	[int])

AS UPDATE [Languages] 

SET  [lngLanguageId]	 = @lngLanguageId,
	 [lngPhraseId]	 = @lngPhraseId 

WHERE 
	( [lngLanguageId]	 = @lngLanguageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Locations_Add]
	(@lngLocationId 	[int] OUTPUT,
	 @strLocation 	[varchar](50))

AS INSERT INTO [Locations] 
	 ( [strLocation]) 
 
VALUES 
	(  @strLocation)

set @lngLocationId = @@identity
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Locations_Delete]
	(@lngLocationId 	[int])

AS DELETE [Locations] 

WHERE 
	( [lngLocationId]	 = @lngLocationId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Locations_Delete_]
	(@lngLocationId 	[int])

AS DELETE [Locations] 

WHERE 
	( [lngLocationId]	 = @lngLocationId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Locations_Update]
	(@lngLocationId 	[int],
	 @strLocation 	[varchar](50))

AS UPDATE  [Locations] 

SET  [strLocation]	 = @strLocation 

WHERE 
	( [lngLocationId]	 = @lngLocationId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Locations_Update_]
	(@lngLocationId 	[int],
	 @strLocation 	[varchar](50))

AS UPDATE [Locations] 

SET  [strLocation]	 = @strLocation 

WHERE 
	( [lngLocationId]	 = @lngLocationId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Photos_Add]
	(@lngPhotoId 	[int] OUTPUT,
	 @strPhoto 	[varchar](2000))

AS INSERT INTO [Photos] 
	 ( [strPhoto]) 
 
VALUES 
	( @strPhoto)

SET @lngPhotoId = @@identity
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Photos_Delete]
	(@lngPhotoId 	[int])

AS DELETE  [Photos] 

WHERE 
	( [lngPhotoId]	 = @lngPhotoId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Photos_Update]
	(@lngPhotoId 	[int],
	 @strPhoto 	[varchar](2000))

AS UPDATE [Photos] 

SET  [strPhoto]	 = @strPhoto 

WHERE 
	( [lngPhotoId]	 = @lngPhotoId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Types_ delete]
	(@lngTypeId 	[int])

AS DELETE.[Types] 

WHERE 
	( [lngTypeId]	 = @lngTypeId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Types_add]
	(@lngTypeId 	[int],
	 @strType	[varchar](50))

AS INSERT INTO [Types] 
	 ( [lngTypeId],
	 [strType]) 
 
VALUES 
	( @lngTypeId,
	 @strType)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Types_update]
	(@lngTypeId 	[int],
	 @strType 	[varchar](50))

AS UPDATE [Types] 

SET  	 [strType]	 = @strType 

WHERE 
	( [lngTypeId]	 = @lngTypeId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Users_add]
	(@lngUserId 	[int],
	 @strUserName 	[varchar](50),
	 @strFirstName 	[varchar](50),
	 @strLastName 	[varchar](50),
	 @strEmail 	[varchar](50),
	 @strPassword 	[varchar](50))

AS INSERT INTO [Users] 
	 ( [lngUserId],
	 [strUserName],
	 [strFirstName],
	 [strLastName],
	 [strEmail],
	 [strPassword]) 
 
VALUES 
	( @lngUserId,
	 @strUserName,
	 @strFirstName,
	 @strLastName,
	 @strEmail,
	 @strPassword)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Users_delete]
	(@lngUserId 	[int])

AS DELETE [Users] 

WHERE 
	( [lngUserId]	 = @lngUserId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Users_update]
	(@lngUserId 	[int],
	 @strUserName 	[varchar](50),
	 @strFirstName 	[varchar](50),
	 @strLastName 	[varchar](50),
	 @strEmail 	[varchar](50),
	 @strPassword 	[varchar](50))

AS UPDATE [Users] 

SET  [lngUserId]	 = @lngUserId,
	 [strUserName]	 = @strUserName,
	 [strFirstName]	 = @strFirstName,
	 [strLastName]	 = @strLastName,
	 [strEmail]	 = @strEmail,
	 [strPassword]	 = @strPassword 

WHERE 
	( [lngUserId]	 = @lngUserId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [CategoriesImages_Add]
	(@lngCategoryId 	[int],
	 @lngImageId 	[int])

AS INSERT INTO [CategoriesImages] 
	 ( [lngCategoryId],
	 [lngImageId]) 
 
VALUES 
	( @lngCategoryId,
	 @lngImageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [CategoriesImages_Delete]
	(@lngCategoryId 	[int],
	 @lngImageId 	[int])

AS DELETE [CategoriesImages] 

WHERE 
	( [lngCategoryId]	 = @lngCategoryId AND
	 [lngImageId]	 = @lngImageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Customers_Add]
	(@lngCustomerId 	[int],
	 @strFirstName 	[varchar](50),
	 @strLastName 	[varchar](50),
	 @strAddress 	[varchar](50),
	 @strCity 	[varchar](50),
	 @strState 	[varchar](50),
	 @strZip 	[varchar](50),
	 @strCountry 	[varchar](50),
	 @strPhone 	[varchar](50),
	 @strFax 	[varchar](50),
	 @strEmail 	[varchar](50),
	 @strOrganizationNr 	[varchar](50),
	 @strWebSite 	[varchar](50),
	 @strNotes 	[varchar](8000),
	 @lngLanguageId 	[int])

AS INSERT INTO [Customers] 
	 ( [lngCustomerId],
	 [strFirstName],
	 [strLastName],
	 [strAddress],
	 [strCity],
	 [strState],
	 [strZip],
	 [strCountry],
	 [strPhone],
	 [strFax],
	 [strEmail],
	 [strOrganizationNr],
	 [strWebSite],
	 [strNotes],
	 [lngLanguageId]) 
 
VALUES 
	( @lngCustomerId,
	 @strFirstName,
	 @strLastName,
	 @strAddress,
	 @strCity,
	 @strState,
	 @strZip,
	 @strCountry,
	 @strPhone,
	 @strFax,
	 @strEmail,
	 @strOrganizationNr,
	 @strWebSite,
	 @strNotes,
	 @lngLanguageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Customers_Delete]
	(@lngCustomerId 	[int])

AS DELETE [Customers] 

WHERE 
	( [lngCustomerId]	 = @lngCustomerId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Customers_Update]
	(@lngCustomerId 	[int],
	 @strFirstName 	[varchar](50),
	 @strLastName 	[varchar](50),
	 @strAddress 	[varchar](50),
	 @strCity 	[varchar](50),
	 @strState 	[varchar](50),
	 @strZip 	[varchar](50),
	 @strCountry 	[varchar](50),
	 @strPhone 	[varchar](50),
	 @strFax 	[varchar](50),
	 @strEmail 	[varchar](50),
	 @strOrganizationNr 	[varchar](50),
	 @strWebSite 	[varchar](50),
	 @strNotes 	[varchar](8000),
	 @lngLanguageId 	[int])

AS UPDATE [Customers] 

SET  [strFirstName]	 = @strFirstName,
	 [strLastName]	 = @strLastName,
	 [strAddress]	 = @strAddress,
	 [strCity]	 = @strCity,
	 [strState]	 = @strState,
	 [strZip]	 = @strZip,
	 [strCountry]	 = @strCountry,
	 [strPhone]	 = @strPhone,
	 [strFax]	 = @strFax,
	 [strEmail]	 = @strEmail,
	 [strOrganizationNr]	 = @strOrganizationNr,
	 [strWebSite]	 = @strWebSite,
	 [strNotes]	 = @strNotes,
	 [lngLanguageId]	 = @lngLanguageId 

WHERE 
	( [lngCustomerId]	 = @lngCustomerId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [FilesPhrases_Fetch] 

@lngLanguageId  [int], @lngFileId [int]

AS


SELECT lngPhraseId, strPhrase FROM [Text] 
WHERE lngLanguageId = @lngLanguageId 
AND lngPhraseId IN(SELECT lngPhraseId FROM FilesPhrases WHERE lngFileId = @lngFileId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [FilesPhrases_add]
	(@lngPhraseId 	[int])

AS INSERT INTO [FilesPhrases] 
	 ( [lngPhraseId]) 
 
VALUES 
	( @lngPhraseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [FilesPhrases_delete]
	(@lngFileId 	[int],
	 @lngPhraseId 	[int])

AS DELETE [FilesPhrases] 

WHERE 
	( [lngFileId]	 = @lngFileId AND
	 [lngPhraseId]	 = @lngPhraseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [FilesPhrases_update]
	(@lngFileId 	[int],
	 @lngPhraseId 	[int])

AS UPDATE [FilesPhrases] 

SET  [lngPhraseId]	 = @lngPhraseId 

WHERE 
	( [lngFileId]	 = @lngFileId AND
	 [lngPhraseId]	 = @lngPhraseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Files_delete]
	(@lngFileId 	[int])
	--@lngCount	[int]=0 output)


AS

--SELECT @lngCount= count(lngFileId) FROM Files WHERE lngFileId = @lngFileId

--IF @lngCount > 0


BEGIN
	DELETE FROM FilesPhrases WHERE lngFileId = @lngFileId
END

BEGIN 
	 DELETE from  [Files] WHERE  ( [lngFileId] = @lngFileId)
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [HousesImages_Add]
	(@lngImageId 	[int],
	 @lngHouseId 	[int])

AS INSERT INTO [HousesImages] 
	 ( [lngImageId],
	 [lngHouseId]) 
 
VALUES 
	( @lngImageId,
	 @lngHouseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [HousesImages_Delete]
	(@lngImageId 	[int],
	 @lngHouseId 	[int])

AS DELETE [HousesImages] 

WHERE 
	( [lngImageId]	 = @lngImageId AND
	 [lngHouseId]	 = @lngHouseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- December 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Images_Fetch] 
@lngCategoryId [int], 
@lngLocationId [int],
@lngHouseId [int]

AS

IF @lngCategoryId <> 0
	BEGIN
		SELECT     Photos.strPhoto, Photos.lngPhotoId, PhotosImages.lngImageId, Images.strImage
		FROM         PhotosImages INNER JOIN
                      	Images ON PhotosImages.lngImageId = Images.lngImageId INNER JOIN
                     	 LocationsImages ON Images.lngImageId = LocationsImages.lngImageId INNER JOIN
                     	 HousesImages ON Images.lngImageId = HousesImages.lngImageId INNER JOIN
                     	 Photos ON PhotosImages.lngPhotoId = Photos.lngPhotoId INNER JOIN
                     	 CategoriesImages ON Images.lngImageId = CategoriesImages.lngImageId
		WHERE     (CategoriesImages.lngCategoryId = @lngCategoryId)
	
	END

IF @lngLocationId <> 0
	BEGIN
		SELECT     Photos.strPhoto, Photos.lngPhotoId, PhotosImages.lngImageId, Images.strImage
		FROM         PhotosImages INNER JOIN
                      	Images ON PhotosImages.lngImageId = Images.lngImageId INNER JOIN
                      	LocationsImages ON Images.lngImageId = LocationsImages.lngImageId INNER JOIN
                      	HousesImages ON Images.lngImageId = HousesImages.lngImageId INNER JOIN
                      	Photos ON PhotosImages.lngPhotoId = Photos.lngPhotoId INNER JOIN
                     	CategoriesImages ON Images.lngImageId = CategoriesImages.lngImageId
		WHERE     (LocationsImages.lngLocationId = @lngLocationId)
	END

IF @lngHouseId <> 0
	BEGIN
		SELECT  Photos.strPhoto, Photos.lngPhotoId, PhotosImages.lngImageId, Images.strImage FROM  PhotosImages INNER JOIN
                      	Images ON PhotosImages.lngImageId = Images.lngImageId INNER JOIN
                      	LocationsImages ON Images.lngImageId = LocationsImages.lngImageId INNER JOIN
                     	 HousesImages ON Images.lngImageId = HousesImages.lngImageId INNER JOIN
                      	Photos ON PhotosImages.lngPhotoId = Photos.lngPhotoId INNER JOIN
                      	CategoriesImages ON Images.lngImageId = CategoriesImages.lngImageId
		WHERE     (HousesImages.lngHouseId = @lngHouseId)	
	END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [LocationsImages_Add]
	(@lngLocationId 	[int],
	 @lngImageId 	[int])

AS INSERT INTO [LocationsImages] 
	 ( [lngLocationId],
	 [lngImageId]) 
 
VALUES 
	( @lngLocationId,
	 @lngImageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [LocationsImages_Delete]
	(@lngLocationId 	[int],
	 @lngImageId 	[int])

AS DELETE [LocationsImages] 

WHERE 
	( [lngLocationId]	 = @lngLocationId AND
	 [lngImageId]	 = @lngImageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [PhotosImages_Add]
	(@lngPhotoId 	[int],
	 @lngImageId 	[int])

AS INSERT INTO [PhotosImages] 
	 ( [lngPhotoId],
	 [lngImageId]) 
 
VALUES 
	( @lngPhotoId,
	 @lngImageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [PhotosImages_Delete]
	(@lngPhotoId 	[int],
	 @lngImageId 	[int])

AS DELETE [PhotosImages] 

WHERE 
	( [lngPhotoId]	 = @lngPhotoId AND
	 [lngImageId]	 = @lngImageId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [PhrasesTypes_add]
	(@lngPhraseTypeId 	[int],
	 @lngPhraseId	[int])

AS INSERT INTO [PhrasesTypes] 
	 ( [lngPhraseTypeId],
	 [lngPhraseId]) 
 
VALUES 
	( @lngPhraseTypeId,
	 @lngPhraseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [PhrasesTypes_delete]
	(@lngPhraseTypeId 	[int],
	 @lngPhraseId 	[int])

AS DELETE [PhrasesTypes] 

WHERE 
	( [lngPhraseTypeId]	 = @lngPhraseTypeId AND
	 [lngPhraseId]	 = @lngPhraseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [PhrasesTypes_update]
	(@lngPhraseTypeId 	[int],
	 @lngPhraseId 	[int])

AS UPDATE [PhrasesTypes] 

SET  [lngPhraseTypeId]	 = @lngPhraseTypeId,
	 [lngPhraseId]	 = @lngPhraseId 

WHERE 
	( [lngPhraseTypeId]	 = @lngPhraseTypeId AND
	 [lngPhraseId]	 = @lngPhraseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Phrases_Fetch] 

@lngLanguageId  [int]

AS


SELECT lngPhraseId, strPhrase FROM [Text] WHERE lngLanguageId = @lngLanguageId
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Phrases_add]
	(@strPhrase	[varchar] (8000),
	@lngLanguageId [int],
	@lngFileId [int],
	@lngPhraseId 	[int] OUTPUT)

AS 
BEGIN
	SELECT @lngPhraseId = MAX(lngPhraseId)+1 FROM Phrases
END

BEGIN
	INSERT INTO [Phrases] ( [lngPhraseId],[strPhrase])  VALUES ( @lngPhraseId,@strPhrase)
END

BEGIN

	INSERT INTO [Text] (lngPhraseId,lngLanguageId,strPhrase) VALUES( @lngPhraseId,@lngLanguageId,@strPhrase)
END

BEGIN
	INSERT INTO FilesPhrases (lngFileId,lngPhraseId) VALUES (@lngFileId,@lngPhraseId)
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Phrases_delete]
	(@lngPhraseId 	[int])

AS


BEGIN
	DELETE [Text]  WHERE [lngPhraseId] = @lngPhraseId

END


BEGIN

	DELETE FilePhrases WHERE lngPhraseId = @lngPhraseId
END



BEGIN
	DELETE [Phrases]  WHERE  [lngPhraseId]  = @lngPhraseId

END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Phrases_update]
	(@lngPhraseId 	[int],
	@lngLanguageId [int],
	 @strPhrase 	[varchar](8000))

AS 

BEGIN
	UPDATE [Text] SET [strPhrase]= @strPhrase WHERE lngPhraseId=@lngPhraseId AND lngLanguageId = @lngLanguageId

END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Text_Fetch] 
	@lngLanguageId [int]
AS

SELECT * FROM [Text] WHERE lngLanguageId = @lngLanguageId
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Text_add]
	(@lngPhraseId 	[int],
	 @lngLanguageId 	[int],
	 @strPhrase 	[varchar](8000))

AS INSERT INTO [Text] 
	 ( [lngPhraseId],
	 [lngLanguageId],
	 [strPhrase]) 
 
VALUES 
	( @lngPhraseId,
	 @lngLanguageId,
	 @strPhrase)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Text_delete]
	(@lngPhraseId 	[int],
	 @lngLanguageId 	[int],
	 @strPhrase 	[varchar])

AS DELETE [Text] 

WHERE 
	( [lngPhraseId]	 = @lngPhraseId AND
	 [lngLanguageId]	 = @lngLanguageId AND
	 [strPhrase]	 = @strPhrase)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
-- SMART DATABASE  
-- Universal language e-commerce library
-- Programmer:	Max Haase         maxhaase@gmail.com
-- November 2000
-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


CREATE PROCEDURE [Text_update]
	(@lngPhraseId 	[int],
	 @lngLanguageId 	[int],
	 @strPhrase 	[varchar](8000))

AS UPDATE [SMART].[dbo].[Text] 

SET  [lngPhraseId]	 = @lngPhraseId,
	 [lngLanguageId]	 = @lngLanguageId,
	 [strPhrase]	 = @strPhrase 

WHERE 
	( [lngPhraseId]	 = @lngPhraseId AND
	 [lngLanguageId]	 = @lngLanguageId AND
	 [strPhrase]	 = @strPhrase)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [Cases_Delete]
	(@lngCaseId 	[int])

AS 

BEGIN
DELETE [Cases] 

WHERE 
	( [lngCaseId]	 = @lngCaseId)
END

BEGIN

DELETE CustomersCases WHERE lngCaseId = @lngCaseId

END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE [Cases_Fetch] 
@lngCustomerId [int]
 AS
SELECT * FROM CASES WHERE lngCaseId IN(SELECT lngCaseId FROM [CustomersCases] WHERE lngCustomerId = @lngCustomerId)


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [CustomersCases_Add]
	(@lngCustomerId 	[int],
	 @lngCaseId 	[int])

AS INSERT INTO [CustomersCases] 
	 ( [lngCustomerId],
	 [lngCaseId]) 
 
VALUES 
	( @lngCustomerId,
	 @lngCaseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [CustomersCases_Delete]
	(@lngCustomerId 	[int],
	 @lngCaseId 	[int])

AS DELETE  [CustomersCases] 

WHERE 
	( [lngCustomerId]	 = @lngCustomerId AND
	 [lngCaseId]	 = @lngCaseId)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

