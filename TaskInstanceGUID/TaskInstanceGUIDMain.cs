using EPDM.Interop.epdm;
using System.Runtime.InteropServices;

namespace TaskInstanceGUID
{
    [ComVisible(true)]
    [Guid("2800F580-B258-48B8-930C-406D60884AFD")]
    public class TaskInstanceGUIDMain : IEdmAddIn5
    {
        #region Public Methods

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            poInfo.mbsAddInName = "Instance GUID Testing";
            poInfo.mbsCompany = "Blue Byte Systems";
            poInfo.mbsDescription = "Testing task instance GUID";
            poInfo.mlAddInVersion = 0;
            poInfo.mlRequiredVersionMinor = 0;
            poInfo.mlRequiredVersionMajor = 10;

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskLaunch);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetupButton);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskDetails);
        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            AttachDebugger();
            var taskProperties = (poCmd.mpoExtra as IEdmTaskProperties);
            switch (poCmd.meCmdType)
            {
                case EdmCmdType.EdmCmd_TaskSetup:
                    taskProperties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsChangeState + (int)EdmTaskFlag.EdmTask_SupportsDetails + (int)EdmTaskFlag.EdmTask_SupportsInitExec + (int)EdmTaskFlag.EdmTask_SupportsScheduling;

                    var cmds = new EdmTaskMenuCmd[1] { new EdmTaskMenuCmd() };
                    cmds[0].mbsMenuString = "Check task GUID...";
                    cmds[0].mbsStatusBarHelp = "Checks the task instance GUID";
                    cmds[0].mlCmdID = 1;
                    cmds[0].mlEdmMenuFlags = (int)EdmMenuFlags.EdmMenu_Nothing;
                    taskProperties.SetMenuCmds(cmds);

                    (poCmd.mpoVault as IEdmVault5).MsgBox(0, $"Add-in name:\n{taskProperties.AddInName}");

                    break;

                case EdmCmdType.EdmCmd_TaskSetupButton:
                    (poCmd.mpoVault as IEdmVault5).MsgBox(0, $"Add-in name:\n{taskProperties.AddInName}");
                    break;

                case EdmCmdType.EdmCmd_TaskDetails:
                    break;

                case EdmCmdType.EdmCmd_TaskRun:
                    break;

                case EdmCmdType.EdmCmd_TaskLaunch:
                    var guid2 = (poCmd.mpoExtra as IEdmTaskInstance).InstanceGUID;
                    (poCmd.mpoVault as IEdmVault5).MsgBox(0, $"Instance GUID:\n{guid2}");

                    var addInName = System.IO.Path.GetFileName(this.GetType().Assembly.Location);
                    (poCmd.mpoVault as IEdmVault5).MsgBox(0, $"Add-in name:\n{addInName}");
                    break;

                case EdmCmdType.EdmCmd_TaskLaunchButton:
                    break;

                default:
                    break;
            }
        }

        public void AttachDebugger()
        {
            System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess();
            System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();

            if (!System.Diagnostics.Debugger.IsAttached)
            {
                System.Diagnostics.Debugger.Launch();
            }
        }

        #endregion
    }
}