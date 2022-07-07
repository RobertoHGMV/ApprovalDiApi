using SAPbobsCOM;
using System.Collections.Generic;

namespace ApprovalDiApi.Services
{
    public class ApprovationSboService
    {
        public void ActiveAndCheckApproval(Company company, Documents document, List<int> wtmCodes)
        {
            SetApprovalTemplateIsActive(company, wtmCodes, BoYesNoEnum.tYES);
            CheckForApproval(document);
        }

        public void DisableApprovalTemplates(Company company, List<int> wtmCodes)
        {
            SetApprovalTemplateIsActive(company, wtmCodes, BoYesNoEnum.tYES);
        }

        private void SetApprovalTemplateIsActive(Company company, List<int> wtmCodes, BoYesNoEnum boYesNoEnum)
        {
            var companyService = company.GetCompanyService();
            var approvalTemplatesService = companyService.GetBusinessService(ServiceTypes.ApprovalTemplatesService) as ApprovalTemplatesService;

            foreach (var code in wtmCodes)
            {
                var oApprovalTemplateParams = approvalTemplatesService?.GetDataInterface(ApprovalTemplatesServiceDataInterfaces.atsdiApprovalTemplateParams) as ApprovalTemplateParams;

                if (oApprovalTemplateParams != null)
                {
                    oApprovalTemplateParams.Code = code;
                    var approvalTemplate = approvalTemplatesService.GetApprovalTemplate(oApprovalTemplateParams);
                    approvalTemplate.IsActive = boYesNoEnum;
                    approvalTemplatesService.UpdateApprovalTemplate(approvalTemplate);
                }
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
