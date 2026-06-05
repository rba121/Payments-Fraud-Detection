--Total payments and flagged transactions summary
SELECT
    COUNT(*) AS total_transactions,
    SUM(amount) AS total_disbursed,
    SUM(is_flagged) AS flagged_count,
    ROUND(100.0 * SUM(is_flagged) / COUNT(*), 2) AS flag_rate_pct,
    ROUND(SUM(CASE WHEN is_flagged = 1 THEN amount ELSE 0 END), 2) AS flagged_amount
FROM transactions;

--Fraud Breakdown: Count and total by flag reason
SELECT
    flag_reason,
    COUNT(*) AS occurrences,
    ROUND(SUM(amount), 2) AS total_amount,
    ROUND(AVG(amount), 2) AS avg_amount
FROM transactions
WHERE is_flagged = 1
GROUP BY flag_reason
ORDER BY total_amount DESC;

--Overpayment Detection: Payments exceeding threshold (> $4500)
SELECT
    transaction_id,
    claimant_id,
    claimant_name,
    payment_type,
    amount,
    payment_date,
    region
FROM transactions
WHERE amount > 4500
ORDER BY amount DESC;

--Duplicate Detection
SELECT
    t1.claimant_id,
    t1.claimant_name,
    t1.payment_date,
    t1.amount,
    COUNT(*) AS duplicate_count
FROM transactions t1
JOIN transactions t2
    ON t1.claimant_id = t2.claimant_id
    AND t1.payment_date = t2.payment_date
    AND t1.transaction_id != t2.transaction_id
GROUP BY t1.claimant_id, t1.payment_date, t1.amount
HAVING COUNT(*) > 1
ORDER BY duplicate_count DESC;

--High Risks Claimants: Multiple flags 
SELECT
    claimant_id,
    claimant_name,
    COUNT(*) AS total_transactions,
    SUM(is_flagged) AS total_flags,
    ROUND(SUM(amount), 2) AS total_paid,
    ROUND(SUM(CASE WHEN is_flagged = 1 THEN amount ELSE 0 END), 2) AS flagged_amount
FROM transactions
GROUP BY claimant_id, claimant_name
HAVING SUM(is_flagged) >= 2
ORDER BY total_flags DESC;

--Regional Breakdown
SELECT
    region,
    COUNT(*) AS total_payments,
    ROUND(SUM(amount), 2) AS total_amount,
    SUM(is_flagged) AS flagged_count,
    ROUND(100.0 * SUM(is_flagged) / COUNT(*), 2) AS flag_rate_pct
FROM transactions
GROUP BY region
ORDER BY flag_rate_pct DESC;

--Monthly Trend
SELECT
    SUBSTR(payment_date, 1, 7) AS month,
    COUNT(*) AS total_payments,
    ROUND(SUM(amount), 2) AS total_disbursed,
    SUM(is_flagged) AS flags
FROM transactions
GROUP BY month
ORDER BY month;

--Payment Type Analysis
SELECT
    payment_type,
    COUNT(*) AS total,
    SUM(is_flagged) AS flagged,
    ROUND(100.0 * SUM(is_flagged) / COUNT(*), 2) AS fraud_rate_pct,
    ROUND(AVG(amount), 2) AS avg_payment
FROM transactions
GROUP BY payment_type
ORDER BY fraud_rate_pct DESC;

--Processor Review:  Flag counts by employee who processed payment
SELECT
    processed_by,
    COUNT(*) AS total_processed,
    SUM(is_flagged) AS flags_generated,
    ROUND(100.0 * SUM(is_flagged) / COUNT(*), 2) AS flag_rate_pct
FROM transactions
GROUP BY processed_by
ORDER BY flag_rate_pct DESC;

--Detailed Flagged transactions summary
SELECT
    transaction_id,
    claimant_id,
    claimant_name,
    payment_type,
    amount,
    payment_date,
    region,
    flag_reason,
    processed_by
FROM transactions
WHERE is_flagged = 1
ORDER BY payment_date DESC;
