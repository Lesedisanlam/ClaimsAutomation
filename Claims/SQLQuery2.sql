CREATE TABLE [dbo].[CollectionMethodData] (
    [Scenario_ID]    INT           NULL,
    [ID]             INT           IDENTITY (1, 1) NOT NULL,
    [collectionmethod] VARCHAR (255) NULL,
    [employee_number1] VARCHAR (255) NULL,
    FOREIGN KEY ([Scenario_ID]) REFERENCES [dbo].[PS_Scenarios] ([ID]) ON DELETE CASCADE

);

CREATE TABLE [dbo].[ChangeLifeData] (

    [Scenario_ID]    INT           NULL,
    [ID]             INT           IDENTITY (1, 1) NOT NULL,
    [Title]          VARCHAR (255) NULL,
    [surname]        VARCHAR (255) NULL,
    [MaritalStatus]  VARCHAR (255) NULL,
    [EducationLevel] VARCHAR (255) NULL,
    [Department]     VARCHAR (255) NULL,
    [Profession]     VARCHAR (255) NULL,
    [Roleplayer]     VARCHAR (255) NULL,
    [RolePlayer_idNum]     VARCHAR (255) NULL,

	FOREIGN KEY ([Scenario_ID]) REFERENCES [dbo].[PS_Scenarios] ([ID]) ON DELETE CASCADE


);
CREATE TABLE [dbo].[ComponentDowngradeUpgrade] (

    [Scenario_ID]    INT           NULL,
    [ID]             INT           IDENTITY (1, 1) NOT NULL,
    [component]         VARCHAR (255) NULL,
    [Method]       VARCHAR (255) NULL,
    [Cover_Amount] VARCHAR (255) NULL,
    [RolePlayer_idNo] VARCHAR (255) NULL,
    

	FOREIGN KEY ([Scenario_ID]) REFERENCES [dbo].[PS_Scenarios] ([ID]) ON DELETE CASCADE
);






