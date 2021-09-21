USE [master]
GO

if DB_ID('db-eng-oc-qe-perf') is not null return

/****** Object:  Database [db-eng-oc-qe-perf]    Script Date: 1/17/2019 2:55:35 PM ******/
CREATE DATABASE [db-eng-oc-qe-perf]
	CONTAINMENT = NONE
	ON  PRIMARY 
( NAME = N'db-eng-oc-qe-perf', FILENAME = N'E:\SQL2K12\Data\db-eng-oc-qe-perf.mdf' , SIZE = 131072KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
	LOG ON 
( NAME = N'db-eng-oc-qe-perf_log', FILENAME = N'F:\SQL2K12\Log\db-eng-oc-qe-perf_log.ldf' , SIZE = 32768KB , MAXSIZE = 2048GB , FILEGROWTH = 32768KB )
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET COMPATIBILITY_LEVEL = 110
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [db-eng-oc-qe-perf].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET ARITHABORT OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET  DISABLE_BROKER 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET RECOVERY FULL 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET  MULTI_USER 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET DB_CHAINING OFF 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET TARGET_RECOVERY_TIME = 0 SECONDS 
GO
EXEC sys.sp_db_vardecimal_storage_format N'db-eng-oc-qe-perf', N'ON'
GO
USE [db-eng-oc-qe-perf]
GO
/****** Object:  Table [dbo].[MACHINE]    Script Date: 1/17/2019 2:55:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[MACHINE](
	[machine_id] [int] IDENTITY(1,1) NOT NULL,
	[machine_name] [nvarchar](128) NOT NULL,
	[version_number] [int] NOT NULL,
	[cpu_speed] [float] NOT NULL,
	[cores] [int] NOT NULL,
	[memory] [int] NOT NULL,
	[operating_system] [nvarchar](128) NOT NULL,
	[office_version] [nvarchar](128) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[machine_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TEST_DEFINITION]    Script Date: 1/17/2019 2:55:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TEST_DEFINITION](
	[test_definition_id] [int] IDENTITY(1,1) NOT NULL,
	[file_identifier] [nvarchar](64) NOT NULL,
	[title] [nvarchar](128) NOT NULL,
	[steps] [nvarchar](2048) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[test_definition_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TEST_RUN]    Script Date: 1/17/2019 2:55:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TEST_RUN](
	[test_run_id] [int] IDENTITY(1,1) NOT NULL,
	[time] [float] NULL,
	[cpu_max_one_sigma] [float] NOT NULL,
	[cpu_max_two_sigma] [float] NOT NULL,
	[cpu_max] [float] NOT NULL,
	[cpu_mean_one_sigma] [float] NOT NULL,
	[cpu_mean_two_sigma] [float] NOT NULL,
	[cpu_mean] [float] NOT NULL,
	[memory_max] [bigint] NULL,
	[memory_net] [bigint] NULL,
	[test_definition_id] [int] NOT NULL,
	[test_suite_run_id] [int] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[test_run_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TEST_SUITE_RUN]    Script Date: 1/17/2019 2:55:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TEST_SUITE_RUN](
	[test_suite_run_id] [int] IDENTITY(1,1) NOT NULL,
	[branch_name] [nvarchar](150) NOT NULL,
	[date_time] [datetime] NULL,
	[machine_id] [int] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[test_suite_run_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[TEST_RUN]  WITH CHECK ADD FOREIGN KEY([test_definition_id])
REFERENCES [dbo].[TEST_DEFINITION] ([test_definition_id])
GO
ALTER TABLE [dbo].[TEST_RUN]  WITH CHECK ADD FOREIGN KEY([test_definition_id])
REFERENCES [dbo].[TEST_DEFINITION] ([test_definition_id])
GO
ALTER TABLE [dbo].[TEST_RUN]  WITH CHECK ADD FOREIGN KEY([test_suite_run_id])
REFERENCES [dbo].[TEST_SUITE_RUN] ([test_suite_run_id])
GO
ALTER TABLE [dbo].[TEST_SUITE_RUN]  WITH CHECK ADD FOREIGN KEY([machine_id])
REFERENCES [dbo].[MACHINE] ([machine_id])
GO
USE [master]
GO
ALTER DATABASE [db-eng-oc-qe-perf] SET  READ_WRITE 
GO
