-- Add DELETE policy to user_notifications table
-- This allows authenticated users to clear or delete their own notifications

DROP POLICY IF EXISTS "Users can delete their own notifications" ON public.user_notifications;

CREATE POLICY "Users can delete their own notifications" ON public.user_notifications
  FOR DELETE TO authenticated USING (auth.uid() = user_id);
