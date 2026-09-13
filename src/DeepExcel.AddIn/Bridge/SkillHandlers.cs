using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DeepExcel.AddIn.Diagnostics;
using DeepExcel.AddIn.Skills;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// Skill library bridge messages.
    /// </summary>
    public partial class MessageBridge
    {
        private readonly SkillStore _skillStore = new SkillStore();

        private static object ToView(Skill skill)
        {
            return new
            {
                id = skill.Id,
                name = skill.Name,
                description = skill.Description,
                run_count = skill.RunCount,
                last_run_at = skill.LastRunAt,
                created_at = skill.CreatedAt,
                step_count = skill.Steps.Count,
                tools = skill.Steps.Select(s => s.Tool).Distinct().ToArray(),
                parameters = skill.Parameters.Select(p => new
                {
                    name = p.Name,
                    kind = p.Kind.ToString().ToLowerInvariant(),
                    label = p.Label,
                    default_value = p.DefaultValue
                })
            };
        }

        private string HandleSkillList()
        {
            try
            {
                return MakeResponse("skill_list", new
                {
                    skills = _skillStore.List().Select(ToView),
                    // Non-null when the last task is worth offering to save.
                    candidate = GetSkillCandidate()
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSkillList failed", ex);
                return MakeError("读取技能列表失败");
            }
        }

        /// <summary>Saves the last successful run as a skill.</summary>
        private string HandleSkillSave(Message msg)
        {
            try
            {
                var name = ReadString(msg, "name");
                if (string.IsNullOrWhiteSpace(name))
                {
                    return MakeError("请给技能起个名字");
                }

                var steps = _lastSuccessfulSteps;
                if (steps == null || steps.Length == 0)
                {
                    // The offer is made right after a task, so this means the
                    // session was restarted or the task was not successful.
                    return MakeError("没有可保存的操作记录，请先成功执行一次任务");
                }

                var skill = SkillParameterizer.Build(name, _lastSuccessfulRequest, steps);
                skill.Description = ReadString(msg, "description");

                if (!_skillStore.Save(skill))
                {
                    return MakeError("技能保存失败");
                }
                Logger.Instance.Info("MessageBridge",
                    $"Skill saved: {skill.Name}, {skill.Steps.Count} steps, {skill.Parameters.Count} params");
                return MakeResponse("skill_saved", ToView(skill));
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSkillSave failed", ex);
                return MakeError("技能保存失败");
            }
        }

        // ------------------------------------------------------------------
        // Cloud sync
        // ------------------------------------------------------------------

        private SkillSyncClient RequireSync()
        {
            if (AccountSession == null || AccountSession.State == Account.SessionState.SignedOut)
            {
                throw new Account.AuthException(
                    Account.AuthFailure.SessionExpired, "技能同步需要先登录账号。");
            }
            return new SkillSyncClient(AccountSession);
        }

        /// <summary>
        /// Uploads every local skill.
        ///
        /// Paths and workbook names are stripped before anything leaves the
        /// machine; the response reports how many skills had something removed
        /// so the user is told rather than surprised.
        /// </summary>
        private string HandleSkillSync()
        {
            try
            {
                var sync = RequireSync();
                var local = _skillStore.List();
                var uploaded = 0;
                var redacted = 0;

                foreach (var skill in local)
                {
                    if (SkillScrubber.WouldRedact(skill))
                    {
                        redacted++;
                    }
                    RunSync(() => sync.UploadAsync(skill));
                    uploaded++;
                }

                var remote = RunSync(() => sync.ListAsync());
                return MakeResponse("skill_synced", new
                {
                    uploaded,
                    redacted,
                    remote_count = remote.Count,
                    remote = remote.Select(r => new
                    {
                        skill_id = r.SkillId,
                        name = r.Name,
                        share_code = r.ShareCode
                    })
                });
            }
            catch (Account.AuthException ex)
            {
                return MakeError(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSkillSync failed", ex);
                return MakeError("技能同步失败");
            }
        }

        /// <summary>Pulls skills that exist on the server but not locally.</summary>
        private string HandleSkillPull()
        {
            try
            {
                var sync = RequireSync();
                var remote = RunSync(() => sync.ListAsync());
                var restored = 0;

                foreach (var summary in remote)
                {
                    if (_skillStore.Get(summary.SkillId) != null)
                    {
                        // Local wins: it still has working defaults that the
                        // uploaded copy had stripped.
                        continue;
                    }
                    var skill = RunSync(() => sync.DownloadAsync(summary.SkillId));
                    if (skill != null && _skillStore.Save(skill))
                    {
                        restored++;
                    }
                }
                return MakeResponse("skill_pulled", new { restored, remote_count = remote.Count });
            }
            catch (Account.AuthException ex)
            {
                return MakeError(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSkillPull failed", ex);
                return MakeError("技能拉取失败");
            }
        }

        private string HandleSkillShare(Message msg)
        {
            var id = ReadString(msg, "id");
            if (string.IsNullOrEmpty(id))
            {
                return MakeError("缺少技能 ID");
            }
            try
            {
                var skill = _skillStore.Get(id);
                if (skill == null)
                {
                    return MakeError("技能不存在");
                }

                var sync = RequireSync();
                // Share implies sync: the server can only publish what it holds.
                RunSync(() => sync.UploadAsync(skill));
                var code = RunSync(() => sync.ShareAsync(id));

                return MakeResponse("skill_shared", new
                {
                    id,
                    share_code = code,
                    redacted = SkillScrubber.WouldRedact(skill)
                });
            }
            catch (Account.AuthException ex)
            {
                return MakeError(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSkillShare failed", ex);
                return MakeError("分享失败");
            }
        }

        private string HandleSkillImport(Message msg)
        {
            var code = ReadString(msg, "share_code");
            if (string.IsNullOrWhiteSpace(code))
            {
                return MakeError("请输入分享码");
            }
            try
            {
                var sync = RequireSync();
                var skill = RunSync(() => sync.ImportAsync(code.Trim()));
                if (skill == null || skill.Steps.Count == 0)
                {
                    return MakeError("分享码无效或技能内容为空");
                }
                if (!_skillStore.Save(skill))
                {
                    return MakeError("保存导入的技能失败");
                }
                return MakeResponse("skill_imported", ToView(skill));
            }
            catch (Account.AuthException ex)
            {
                return MakeError(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSkillImport failed", ex);
                return MakeError("导入失败");
            }
        }

        private string HandleSkillDelete(Message msg)
        {
            var id = ReadString(msg, "id");
            if (string.IsNullOrEmpty(id))
            {
                return MakeError("缺少技能 ID");
            }
            return _skillStore.Delete(id)
                ? MakeResponse("skill_deleted", new { id })
                : MakeError("技能不存在");
        }

        /// <summary>
        /// Expands a skill into the instruction that will be sent to the agent.
        ///
        /// The panel then submits it as a normal message, so a replay is
        /// indistinguishable from the user typing the request again -- including
        /// going through the same confirmation and preview flow. A replay path
        /// that skipped those checks would turn the skill library into a way to
        /// bypass them.
        /// </summary>
        private string HandleSkillPrepare(Message msg)
        {
            try
            {
                var id = ReadString(msg, "id");
                var skill = _skillStore.Get(id);
                if (skill == null)
                {
                    return MakeError("技能不存在");
                }

                var arguments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (msg?.Payload != null &&
                    msg.Payload.Value.TryGetProperty("arguments", out var supplied) &&
                    supplied.ValueKind == JsonValueKind.Object)
                {
                    foreach (var property in supplied.EnumerateObject())
                    {
                        if (property.Value.ValueKind == JsonValueKind.String)
                        {
                            arguments[property.Name] = property.Value.GetString();
                        }
                    }
                }

                _skillStore.RecordRun(id);
                return MakeResponse("skill_prompt", new
                {
                    id = skill.Id,
                    name = skill.Name,
                    prompt = skill.ToPrompt(arguments)
                });
            }
            catch (Exception ex)
            {
                Logger.Instance.Error("MessageBridge", "HandleSkillPrepare failed", ex);
                return MakeError("技能准备失败");
            }
        }
    }
}
