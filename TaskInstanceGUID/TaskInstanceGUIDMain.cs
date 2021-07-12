using BlueByte.SOLIDWORKS.PDMProfessional.PDMAddInFramework;
using BlueByte.SOLIDWORKS.PDMProfessional.PDMAddInFramework.Attributes;
using EPDM.Interop.epdm;
using System.Runtime.InteropServices;

namespace TaskInstanceGUID
{
    [Name("Instance GUID Testing")]
    [CompanyName("Blue Byte Systems")]
    [Description("Testing task instance GUID")]
    [AddInVersion(false, 0)]
    [RequiredVersion(10, 0)]
    [IsTask(true)]
    [TaskFlags((int)EdmTaskFlag.EdmTask_SupportsChangeState + (int)EdmTaskFlag.EdmTask_SupportsDetails + (int)EdmTaskFlag.EdmTask_SupportsInitExec + (int)EdmTaskFlag.EdmTask_SupportsScheduling)]
    [ComVisible(true)]
    [Guid("2800F580-B258-48B8-930C-406D60884AFD")]
    public class TaskInstanceGUIDMain : AddInBase
    {
        #region Public Methods

        public override void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            base.OnCmd(ref poCmd, ref ppoData);

            AttachDebugger();

            switch (poCmd.meCmdType)
            {
                case EdmCmdType.EdmCmd_TaskSetup:
                    AddContextMenu("Check task GUID...", "It displays current task GUID");
                    break;

                case EdmCmdType.EdmCmd_TaskSetupButton:
                    break;

                case EdmCmdType.EdmCmd_TaskDetails:
                    break;

                case EdmCmdType.EdmCmd_TaskRun:
                    break;

                case EdmCmdType.EdmCmd_TaskLaunch:
                    var guid1 = Instance.InstanceGUID;
                    var guid2 = (poCmd.mpoExtra as IEdmTaskInstance).InstanceGUID;
                    Vault.MsgBox(0, $"Instance.InstanceGUID:\n{guid1}\n\n(poCmd.mpoExtra as IEdmTaskInstance).InstanceGUID:\n{guid2}");
                    break;

                case EdmCmdType.EdmCmd_TaskLaunchButton:
                    break;

                default:
                    break;
            }
        }

        #endregion
    }
}