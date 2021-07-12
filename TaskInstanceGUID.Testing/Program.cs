using BlueByte.SOLIDWORKS.PDMProfessional.UnitTesting;
using BlueByte.SOLIDWORKS.PDMProfessional.UnitTesting.Attributes;
using BlueByte.SOLIDWORKS.PDMProfessional.UnitTesting.Mockers;
using EPDM.Interop.epdm;
using System;

namespace TaskInstanceGUID.Testing
{
    internal class Program
    {
        #region Private Methods

        private static void Main(string[] args)
        {
            var test = new TestClass();
            test.Run();
            Console.ReadLine();
        }

        #endregion
    }

    [TestVault(Name = "BlueByte")]
    public class TestClass : TestableAddInBase<TaskInstanceGUIDMain>
    {
        #region Private Fields

        private string File = @"API\Excel file.xlsx";
        private IEdmTaskInstance taskInstance;
        private IEdmTaskProperties taskProperties;

        #endregion

        #region Public Methods

        public override void Startup()
        {
            base.Startup();
            taskInstance = PDMMockFactory.MockEdmTaskInstance(Vault, "GUID Test Task");
            taskProperties = PDMMockFactory.MockTaskProperties(Vault, "GUID Test Task");
        }

        [PDMTestMethod()]
        public void TestTaskSetup()
        {
            var poCmd = MockEdmCmd.EdmCmd_TaskSetup.Mock(taskProperties);
            var poData = MockEdmCmdData.EdmCmd_TaskSetup.Mock();
            EdmCmdData[] edmCmd = new EdmCmdData[] { poData };
            AddInObject.OnCmd(ref poCmd, ref edmCmd);
        }

        [PDMTestMethod()]
        public void TestTaskSetupButton()
        {
            var poCmd = MockEdmCmd.EdmCmd_TaskSetupButton.Mock(taskProperties, "");
            var poData = MockEdmCmdData.EdmCmd_TaskSetupButton.Mock();
            EdmCmdData[] edmCmd = new EdmCmdData[] { poData };
            AddInObject.OnCmd(ref poCmd, ref edmCmd);
        }

        [PDMTestMethod()]
        public void TestTaskLaunch()
        {
            var object1 = PDMFactory.UseThisVault(Vault).ThenGetFileByRelativePath(File).As<IEdmFile5>();
            var poCmd = MockEdmCmd.EdmCmd_TaskLaunch.Mock(taskInstance);
            var poData = MockEdmCmdData.EdmCmd_TaskLaunch.Mock(object1);
            EdmCmdData[] edmCmd = new EdmCmdData[] { poData };
            AddInObject.OnCmd(ref poCmd, ref edmCmd);
        }

        [PDMTestMethod()]
        public void TestTaskRun()
        {
            var object1 = PDMFactory.UseThisVault(Vault).ThenGetFileByRelativePath(File).As<IEdmFile5>();
            var poCmd = MockEdmCmd.EdmCmd_TaskRun.Mock(taskInstance);
            var poData = MockEdmCmdData.EdmCmd_TaskRun.Mock(object1);
            EdmCmdData[] edmCmd = new EdmCmdData[] { poData };
            AddInObject.OnCmd(ref poCmd, ref edmCmd);
        }



        #endregion
    }
}