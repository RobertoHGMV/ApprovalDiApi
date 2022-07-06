using SAPbobsCOM;

namespace ApprovalDiApi.Services
{
    public class ApprovationSboService
    {
        private ApprovalTemplatesService ApprovalTemplatesService { get; set; }
        private List<ApprovalTemplate> ApprovalTemplatesList { get; set; }

        public ApprovationSboService(Company company)
        {
            ApprovalTemplatesList = new List<ApprovalTemplate>();

            var companyService = company.GetCompanyService();
            ApprovalTemplatesService = companyService.GetBusinessService(ServiceTypes.ApprovalTemplatesService) as ApprovalTemplatesService;
        }

        public void ActiveAndCheckApproval(Documents document, List<int> wtmCodes)
        {
            ActiveApprovalTemplates(wtmCodes);
            CheckForApproval(document);
        }

        private void ActiveApprovalTemplates(List<int> wtmCodes)
        {
            foreach (var code in wtmCodes)
            {
                var oApprovalTemplateParams = ApprovalTemplatesService?.GetDataInterface(ApprovalTemplatesServiceDataInterfaces.atsdiApprovalTemplateParams) as ApprovalTemplateParams;

                if (oApprovalTemplateParams != null)
                {
                    oApprovalTemplateParams.Code = code;
                    var approvalTemplate = ApprovalTemplatesService.GetApprovalTemplate(oApprovalTemplateParams);
                    approvalTemplate.IsActive = BoYesNoEnum.tYES;
                    ApprovalTemplatesService.UpdateApprovalTemplate(approvalTemplate);
                    ApprovalTemplatesList.Add(approvalTemplate);
                }
            }
        }

        public void DisableApprovalTemplates()
        {
            foreach (var approval in ApprovalTemplatesList)
            {
                approval.IsActive = BoYesNoEnum.tNO;
                ApprovalTemplatesService.UpdateApprovalTemplate(approval);
            }
        }

        private void CheckForApproval(Documents document)
        {
            var returnCode = document.GetApprovalTemplates();

            if (returnCode == 0 && document.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                for (var i = 1; i <= document.Document_ApprovalRequests.Count; i++)
                {
                    document.Document_ApprovalRequests.SetCurrentLine(i - 1);
                    document.Document_ApprovalRequests.Remarks = $"Referente a teste de aprovação pelo SAP";
                }
        }
    }
}
