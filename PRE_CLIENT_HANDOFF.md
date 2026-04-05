# Pre-Client Handoff Checklist

Use this checklist before showing Golden Leaf Agency HQ to a client or rolling it out to a live agency team.

## Handoff Summary

Complete this quick score before delivery:

- Hosting works
- Admin flow works
- Producer flow works
- Reporting works
- Claude works
- Mobile is acceptable
- Brand looks client-ready

If any of those answers is "not yet," fix that before live rollout.

## 1. Environment and Access

- Confirm the live Render site opens correctly on desktop and mobile.
- Confirm admin login works.
- Confirm at least one producer login works.
- Confirm the `Help Center` tab appears.
- Confirm Claude opens and answers a basic question.
- Confirm logging out and back in works cleanly.

## 2. Admin Workflow QA

- Create a lead manually as admin.
- Import at least one CSV batch as admin.
- Confirm routing behaves the way the agency expects.
- Reassign a lead to a different producer.
- Add a new producer and confirm the producer appears in the roster.
- Add and remove a carrier from the carrier table.
- Edit assumptions and verify the table locks again after editing.
- Export the workbook and confirm all expected tabs are present.

## 3. Producer Workflow QA

- Log in as a producer.
- Create a self-generated lead and confirm it stays assigned to that producer.
- Update the lead in `Update Leads`.
- Move the lead in `Move Stages`.
- Log an activity on the lead.
- Upload and open an attachment.
- Confirm the producer only sees producer-level commission, not agency economics.

## 4. Renewal Workflow QA

- Mark a lead as `Bound`.
- Add an effective date.
- Confirm renewal tracking picks it up even if policy type started as `New`.
- Confirm expiration date is inferred when appropriate.
- Confirm renewal queue appears on the dashboard.

## 5. Dashboard and Reporting QA

- Confirm dashboard metrics load.
- Confirm action queue shows follow-up priorities.
- Confirm stale alerts appear for overdue or aging leads.
- Confirm owner reports display producer performance and source profitability.
- Confirm the numbers shown in reporting match the underlying lead records.

## 6. Help Center and Claude QA

- Open `Help Center` as admin and review the setup guidance.
- Open `Help Center` as producer and review the workflow guidance.
- Ask Claude a setup question from Help Center.
- Ask Claude a lead question from a producer account.
- Ask Claude an owner-style question from the admin account.
- Confirm Claude respects role scope.

## 7. Integrations Readiness

These can stay unconfigured for launch if the client is not ready yet.

- Google Calendar: verify the page explains setup clearly.
- Email reminders: verify the page explains setup clearly.
- Text reminders: verify the page explains setup clearly.
- If they are not being turned on at launch, make sure the client knows they are optional future upgrades.

## 8. Copy and Presentation Review

- Make sure agency name, logo, and color palette match the client brand.
- Check that no rough internal phrasing remains.
- Check empty states and error states for clarity.
- Check buttons, tabs, and fields for awkward wrapping.
- Confirm the language says `producer` or `admin` where appropriate and avoids developer jargon.

## 9. Mobile Review

- Log in on a phone-sized viewport.
- Check dashboard readability.
- Check pipeline tabs and lead workspace.
- Check Claude panel usability.
- Check that nothing critical is cut off or unusably cramped.

## 10. Demo Readiness

- Create a small clean sample data set.
- Include at least:
  - one fresh lead
  - one quoted lead
  - one overdue follow-up
  - one bound policy
  - one renewal coming due
- Keep names and businesses client-safe and presentation-friendly.
- If you want a faster starting point, use the built-in `Download Demo Leads` button in `Setup > Routing > Bulk Lead Import`.

## Suggested Demo Story

1. Show the owner dashboard first.
2. Show reports and explain what the owner learns in 30 seconds.
3. Create a lead.
4. Update a lead.
5. Move a lead on the board.
6. Upload an attachment.
7. Show Help Center.
8. Ask Claude a real business question.

## Suggested Admin Launch Steps

Give this order to the client admin:

1. Add admins and producers in `Setup > Users`.
2. Review `Setup > Routing` and decide whether owner-fed leads should auto-assign.
3. Review `Setup > Assumptions`.
4. Review `Setup > Carrier Table`.
5. Import or enter the first batch of leads.
6. Confirm dashboards and reports look right.
7. Decide whether to turn on optional integrations later.

## Final Go / No-Go Question

Before delivery, ask:

"If this client used the app tomorrow morning, would the owner know what to review first and would the producers know exactly how to work their day inside it?"

If the answer is yes, the product is ready for a live pilot.
