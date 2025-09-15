CREATE TABLE [dbo].[AdditionalTasks](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ScheduledTaskID] [int] NOT NULL,
	[TaskStart] [datetime2](2) NOT NULL,
	[Config] [xml] NULL,
	[AddedToScheduleOn] [datetime2](2) NULL,
	[CreatedOn] [datetime2](2) NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
	[ModifiedOn] [datetime2](2) NOT NULL,
	[ModifiedBy] [varchar](50) NOT NULL,
 CONSTRAINT [PK_AdditionalTasks] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[GeneralLog](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[LogSource] [varchar](100) NOT NULL,
	[LogStatus] [varchar](50) NOT NULL,
	[LogMessage] [varchar](500) NOT NULL,
	[AddnlInfo] [varchar](500) NULL,
	[CreatedOn] [datetime2](2) NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
 CONSTRAINT [pk_GeneralLog] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ProcedureErrors](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ErrorCode] [int] NULL,
	[ErrorSource] [varchar](200) NULL,
	[AdditionalInformation] [varchar](200) NULL,
	[ErrorDescription] [varchar](200) NULL,
	[CreatedOn] [datetime2](2) NULL,
	[CreatedBy] [varchar](50) NULL,
 CONSTRAINT [PK_ProcedureErrors] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ReportSettings](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ScheduledTaskID] [int] NOT NULL,
	[TemplateLocation] [varchar](max) NULL,
	[EmailSubject] [varchar](500) NULL,
	[EmailBody] [varchar](1500) NULL,
	[EmailToList] [varchar](500) NULL,
	[EmailccList] [varchar](500) NULL,
	[SavePath] [varchar](max) NULL,
	[CreatedOn] [datetime2](2) NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
	[ModifiedOn] [datetime2](2) NOT NULL,
	[ModifiedBy] [varchar](50) NOT NULL,
 CONSTRAINT [PK_ReportSettings_1] PRIMARY KEY CLUSTERED 
(
	[ScheduledTaskID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[TaskResults](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ScheduledTaskID] [int] NOT NULL,
	[TaskScheduleID] [int] NULL,
	[Config] [xml] NULL,
	[TaskExpected] [datetime2](2) NOT NULL,
	[TaskStart] [datetime2](2) NULL,
	[TaskEnd] [datetime2](2) NULL,
	[CompletionMessage] [varchar](500) NULL,
	[TaskCompleted] [bit] NULL,
	[CreatedOn] [datetime2](2) NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
	[ModifiedOn] [datetime2](2) NOT NULL,
	[ModifiedBy] [varchar](50) NOT NULL,
 CONSTRAINT [PK_TaskResults] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Config]    Script Date: 15/09/2025 14:12:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Config](
	[CfgCode] [varchar](10) NOT NULL,
	[CfgDescr] [varchar](200) NOT NULL,
	[CfgValue] [varchar](100) NOT NULL,
 CONSTRAINT [pk_Config] PRIMARY KEY CLUSTERED 
(
	[CfgCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TaskSchedules]    Script Date: 15/09/2025 14:12:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TaskSchedules](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Descr] [varchar](100) NULL,
	[PeriodType] [tinyint] NULL,
	[NthPerPeriod] [tinyint] NULL,
	[StartAt] [time](2) NULL,
	[RecurringTo] [time](2) NULL,
	[RecurringInterval] [smallint] NULL,
	[WorkingDay] [bit] NULL,
	[IsMonday] [bit] NULL,
	[IsTuesday] [bit] NULL,
	[IsWednesday] [bit] NULL,
	[IsThursday] [bit] NULL,
	[IsFriday] [bit] NULL,
	[IsSaturday] [bit] NULL,
	[IsSunday] [bit] NULL,
	[IsRecurring]  AS (case when [RecurringInterval]>=(15) AND [RecurringTo] IS NOT NULL then (1) else (0) end),
	[CreatedOn] [datetime2](2) NULL,
	[CreatedBy] [varchar](50) NULL,
	[ModifiedOn] [datetime2](2) NULL,
	[ModifiedBy] [varchar](50) NULL,
 CONSTRAINT [PK_TaskSchedules] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ScheduledTasks]    Script Date: 15/09/2025 14:12:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ScheduledTasks](
	[ID] [int] IDENTITY(1000,1) NOT NULL,
	[TaskName] [varchar](200) NOT NULL,
	[VersionNo] [varchar](20) NOT NULL,
	[PriorityVal] [tinyint] NOT NULL,
	[TaskScheduleID] [int] NULL,
	[SourceFolder] [varchar](200) NULL,
	[SourceFilename] [varchar](100) NOT NULL,
	[SourceFunction] [varchar](200) NULL,
	[SendTaskID] [bit] NOT NULL,
	[Config] [xml] NULL,
	[IsReport] [bit] NOT NULL,
	[IsActive]  AS (case when [DisabledOn] IS NULL then (1) else (0) end),
	[DisabledOn] [datetime2](2) NULL,
	[DisabledBy] [varchar](50) NULL,
	[CreatedOn] [datetime2](2) NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
	[ModifiedOn] [datetime2](2) NOT NULL,
	[ModifiedBy] [varchar](50) NOT NULL,
 CONSTRAINT [PK_ScheduledTasks] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[ScheduledTasks_TaskSchedules](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ScheduledTaskID] [int] NOT NULL,
	[TaskScheduleID] [int] NOT NULL,
	[IsActive]  AS (case when [DisabledOn] IS NULL then (1) else (0) end),
	[DisabledOn] [datetime2](2) NULL,
	[DisabledBy] [varchar](50) NULL,
	[CreatedOn] [datetime2](2) NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
	[ModifiedOn] [datetime2](2) NOT NULL,
	[ModifiedBy] [varchar](50) NOT NULL,
 CONSTRAINT [PK_ScheduledTasks_TaskSchedules_1] PRIMARY KEY CLUSTERED 
(
	[ScheduledTaskID] ASC,
	[TaskScheduleID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO