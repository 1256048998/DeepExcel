using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DeepExcel.AddIn.Bridge;

namespace DeepExcel.AddIn.Security
{
    /// <summary>
    /// 安全操作网关 - 拦截高风险操作并要求二次验证
    /// </summary>
    public class SecurityGateway
    {
        private readonly SecurityManager _securityManager;
        private readonly Dictionary<string, VerificationState> _pendingVerifications = new();

        public event Func<string, string, Task<bool>> OnVerificationRequired;

        public SecurityGateway(SecurityManager securityManager)
        {
            _securityManager = securityManager;
        }

        /// <summary>
        /// 检查操作是否需要验证
        /// </summary>
        public bool RequiresVerification(string toolName)
        {
            return HighRiskTools.Contains(toolName);
        }

        /// <summary>
        /// 发起二次验证请求
        /// </summary>
        public async Task<VerificationResult> RequestVerification(string toolName, Dictionary<string, object> arguments)
        {
            var code = _securityManager.GenerateVerificationCode();
            var id = Guid.NewGuid().ToString();
            var state = new VerificationState
            {
                Id = id,
                ToolName = toolName,
                Arguments = arguments,
                VerificationCode = code,
                ExpiresAt = DateTime.Now.AddMinutes(5),
                Attempts = 0
            };

            _pendingVerifications[id] = state;

            // 推送验证请求到UI
            bool approved = false;
            if (OnVerificationRequired != null)
            {
                approved = await OnVerificationRequired.Invoke(id, code);
            }

            if (approved)
            {
                _pendingVerifications.Remove(id);
                return new VerificationResult { Approved = true, VerificationId = id };
            }

            return new VerificationResult { Approved = false, VerificationId = id };
        }

        /// <summary>
        /// 验证用户输入的验证码
        /// </summary>
        public bool VerifyCode(string verificationId, string code)
        {
            if (!_pendingVerifications.ContainsKey(verificationId))
                return false;

            var state = _pendingVerifications[verificationId];

            if (state.ExpiresAt < DateTime.Now)
            {
                _pendingVerifications.Remove(verificationId);
                return false;
            }

            state.Attempts++;
            if (state.Attempts > 3)
            {
                _pendingVerifications.Remove(verificationId);
                return false;
            }

            if (state.VerificationCode.Equals(code, StringComparison.OrdinalIgnoreCase))
            {
                _pendingVerifications.Remove(verificationId);
                return true;
            }

            return false;
        }

        /// <summary>
        /// 检查验证状态
        /// </summary>
        public VerificationState GetVerificationState(string verificationId)
        {
            if (_pendingVerifications.ContainsKey(verificationId))
                return _pendingVerifications[verificationId];
            return null;
        }

        /// <summary>
        /// 安全执行工具（自动处理验证）
        /// </summary>
        public async Task<ToolResult> ExecuteWithSecurityCheck(
            string toolName,
            Dictionary<string, object> arguments,
            Func<string, Dictionary<string, object>, Task<ToolResult>> executeFunc)
        {
            if (RequiresVerification(toolName))
            {
                var result = await RequestVerification(toolName, arguments);
                if (!result.Approved)
                {
                    return new ToolResult
                    {
                        Name = toolName,
                        Success = false,
                        Error = "用户取消了高风险操作"
                    };
                }
            }

            return await executeFunc(toolName, arguments);
        }

        // 高风险工具列表
        private static readonly HashSet<string> HighRiskTools = new(StringComparer.OrdinalIgnoreCase)
        {
            "execute_vba",
            "execute_python",
            "remove_duplicates",
            "rollback",
            "clean_data"
        };
    }

    public class VerificationState
    {
        public string Id { get; set; }
        public string ToolName { get; set; }
        public Dictionary<string, object> Arguments { get; set; }
        public string VerificationCode { get; set; }
        public DateTime ExpiresAt { get; set; }
        public int Attempts { get; set; }
    }

    public class VerificationResult
    {
        public bool Approved { get; set; }
        public string VerificationId { get; set; }
    }
}
