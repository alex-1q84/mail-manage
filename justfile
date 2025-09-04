# 增量清理大邮件，清理后大邮件将被移到垃圾箱内
incrmt_tidy_mail:
    python3 check_large_attachments.py

# 全量清理大邮件，清理后大邮件将被移到垃圾箱内
full_tidy_mail:
    python3 check_large_attachments.py --all

# 增量清理大邮件，大邮件会被永久删除
incrmt_tidy_mail_hard_del:
    python3 check_large_attachments.py

# 全量清理大邮件，大邮件会被永久删除
full_tidy_mail_hard_del:
    python3 check_large_attachments.py --all
