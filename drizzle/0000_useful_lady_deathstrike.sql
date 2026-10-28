CREATE TABLE `unsubscribe_attempts` (
	`key` text PRIMARY KEY NOT NULL,
	`status` text NOT NULL,
	`reason` text NOT NULL,
	`expires` integer NOT NULL
);
--> statement-breakpoint
CREATE INDEX `attempt_expiry` ON `unsubscribe_attempts` (`expires`);--> statement-breakpoint
CREATE TABLE `unsubscribe_budgets` (
	`key` text PRIMARY KEY NOT NULL,
	`remaining` integer NOT NULL,
	`expires` integer NOT NULL,
	CONSTRAINT "budget_nonnegative" CHECK("unsubscribe_budgets"."remaining" >= 0)
);
--> statement-breakpoint
CREATE INDEX `budget_expiry` ON `unsubscribe_budgets` (`expires`);