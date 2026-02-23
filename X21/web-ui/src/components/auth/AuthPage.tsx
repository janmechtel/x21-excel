import React, { useState } from "react";
import { supabase } from "../../lib/supabaseClient";
import { Card, CardContent } from "../ui/card";
import { Button } from "../ui/button";
import { Alert, AlertDescription } from "../ui/alert";

export const AuthPage: React.FC = () => {
  const [step, setStep] = useState<"email" | "details" | "code">("email");
  const [email, setEmail] = useState("");
  const [firstName, setFirstName] = useState("");
  const [lastName, setLastName] = useState("");
  const [code, setCode] = useState("");
  const [loading, setLoading] = useState(false);
  const [isReturningUser, setIsReturningUser] = useState(false);
  const [message, setMessage] = useState<{
    type: "error" | "success";
    text: string;
  } | null>(null);

  // Check if user exists and handle accordingly
  const handleEmailNext = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!email.trim()) {
      setMessage({
        type: "error",
        text: "Please enter a valid email address",
      });
      return;
    }

    setLoading(true);
    setMessage(null);

    try {
      if (!supabase) {
        throw new Error("Authentication is not configured");
      }

      console.log("🔍 Checking if user exists:", email);

      // Try to send OTP with shouldCreateUser: false
      // This will succeed if user exists, fail if they don't
      const { error } = await supabase.auth.signInWithOtp({
        email: email.trim(),
        options: {
          shouldCreateUser: false, // Don't create user - just check if they exist
        },
      });

      if (error) {
        // User doesn't exist - collect their details
        console.log("👤 New user detected - collecting details");
        setIsReturningUser(false);
        setStep("details");
      } else {
        // User exists - OTP sent successfully
        console.log("🔄 Returning user - OTP sent directly");
        setIsReturningUser(true);
        setMessage({
          type: "success",
          text: `✅ Welcome back! OTP sent to ${email}. Check your email for a 6-digit code.`,
        });
        setStep("code");
      }
    } catch (error: any) {
      console.error("❌ Error checking user:", error);
      setMessage({
        type: "error",
        text: error.message || "Failed to process email",
      });
    } finally {
      setLoading(false);
    }
  };

  // Send OTP for new users (after collecting details)
  const sendOTPForNewUser = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();

    if (!firstName.trim() || !lastName.trim()) {
      setMessage({
        type: "error",
        text: "Please enter your first and last name",
      });
      return;
    }

    setLoading(true);
    setMessage(null);

    console.log("🔄 Creating new user account:", email, firstName, lastName);

    try {
      if (!supabase) {
        throw new Error("Authentication is not configured");
      }

      const { error } = await supabase.auth.signInWithOtp({
        email: email.trim(),
        options: {
          shouldCreateUser: true,
          data: {
            first_name: firstName.trim(),
            last_name: lastName.trim(),
            full_name: `${firstName.trim()} ${lastName.trim()}`,
          },
        },
      });

      console.log("📧 Supabase new user response:", { error });

      if (error) {
        console.error("❌ Supabase error:", error);
        throw error;
      }

      console.log("✅ New user OTP sent to:", email);

      setMessage({
        type: "success",
        text: `✅ Welcome to X21! Verification code sent to ${email}. Check your email for a 6-digit code.`,
      });
      setStep("code");
    } catch (error: any) {
      console.error("❌ New user creation failed:", error);
      setMessage({
        type: "error",
        text: error.message || "Failed to create account",
      });
    } finally {
      setLoading(false);
    }
  };

  const verifyOTPCode = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setMessage(null);

    try {
      if (!supabase) {
        throw new Error("Authentication is not configured");
      }

      const { data, error } = await supabase.auth.verifyOtp({
        email,
        token: code,
        type: "email",
      });

      if (error) throw error;

      if (data.session) {
        console.log("✅ Successfully authenticated! Session:", data.session);
        // The AuthContext will automatically pick up the session change
      }
    } catch (error: any) {
      setMessage({
        type: "error",
        text: error.message || "Invalid code. Please try again.",
      });
    } finally {
      setLoading(false);
    }
  };

  const goBackToEmail = () => {
    setStep("email");
    setFirstName("");
    setLastName("");
    setCode("");
    setIsReturningUser(false);
    setMessage(null);
  };

  const goBackToDetails = () => {
    setStep("details");
    setCode("");
    setMessage(null);
  };

  // Resend OTP based on user type
  const resendOTP = async () => {
    if (isReturningUser) {
      // For returning users, just resend OTP
      await handleEmailNext(new Event("submit") as any);
    } else {
      // For new users, resend with their details
      await sendOTPForNewUser();
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-50 to-slate-100 dark:from-slate-900 dark:to-slate-800 p-4">
      <Card className="w-full max-w-md shadow-lg">
        <CardContent className="p-6">
          <div className="text-center mb-6">
            <h1 className="text-2xl font-bold text-slate-900 dark:text-slate-100 mb-2">
              Welcome to X21
            </h1>
            <p className="text-sm text-slate-600 dark:text-slate-400">
              {step === "email"
                ? "Enter your email to get started"
                : step === "details"
                ? "Tell us your name to create your account"
                : `Enter the 6-digit code sent to ${email}`}
            </p>
          </div>

          {message && (
            <Alert
              className={`mb-4 ${
                message.type === "error"
                  ? "border-red-200 bg-red-50"
                  : "border-green-200 bg-green-50"
              }`}
            >
              <AlertDescription
                className={
                  message.type === "error" ? "text-red-800" : "text-green-800"
                }
              >
                {message.text}
              </AlertDescription>
            </Alert>
          )}

          {step === "email" ? (
            // Step 1: Email input
            <form onSubmit={handleEmailNext} className="space-y-4">
              <div>
                <label
                  htmlFor="email"
                  className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1"
                >
                  Email address
                </label>
                <input
                  id="email"
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="your.email@example.com"
                  required
                  className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                  disabled={loading}
                />
              </div>

              <Button
                type="submit"
                disabled={loading || !email.trim()}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white font-medium py-2 px-4 rounded-lg transition-colors"
              >
                {loading ? (
                  <div className="flex items-center justify-center">
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin mr-2"></div>
                    Checking...
                  </div>
                ) : (
                  "Continue"
                )}
              </Button>
            </form>
          ) : step === "details" ? (
            // Step 2: Name details (only for new users)
            <form onSubmit={sendOTPForNewUser} className="space-y-4">
              <div>
                <label
                  htmlFor="email-display"
                  className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1"
                >
                  Email address
                </label>
                <input
                  id="email-display"
                  type="email"
                  value={email}
                  disabled
                  className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-400"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label
                    htmlFor="firstName"
                    className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1"
                  >
                    First name
                  </label>
                  <input
                    id="firstName"
                    type="text"
                    value={firstName}
                    onChange={(e) => setFirstName(e.target.value)}
                    placeholder="John"
                    required
                    className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                    disabled={loading}
                  />
                </div>

                <div>
                  <label
                    htmlFor="lastName"
                    className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1"
                  >
                    Last name
                  </label>
                  <input
                    id="lastName"
                    type="text"
                    value={lastName}
                    onChange={(e) => setLastName(e.target.value)}
                    placeholder="Doe"
                    required
                    className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                    disabled={loading}
                  />
                </div>
              </div>

              <Button
                type="submit"
                disabled={loading || !firstName.trim() || !lastName.trim()}
                className="w-full bg-green-600 hover:bg-green-700 text-white font-medium py-2 px-4 rounded-lg transition-colors"
              >
                {loading ? (
                  <div className="flex items-center justify-center">
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin mr-2"></div>
                    Creating account...
                  </div>
                ) : (
                  "Create Account"
                )}
              </Button>

              <Button
                type="button"
                onClick={goBackToEmail}
                variant="outline"
                className="w-full border-slate-300 text-slate-700 hover:bg-slate-50"
                disabled={loading}
              >
                Back
              </Button>
            </form>
          ) : (
            // Step 3: Code verification
            <form onSubmit={verifyOTPCode} className="space-y-4">
              <div>
                <label
                  htmlFor="email-display"
                  className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1"
                >
                  Email address
                </label>
                <input
                  id="email-display"
                  type="email"
                  value={email}
                  disabled
                  className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-400"
                />
              </div>

              <div>
                <label
                  htmlFor="code"
                  className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1"
                >
                  6-digit verification code
                </label>
                <input
                  id="code"
                  type="text"
                  value={code}
                  onChange={(e) =>
                    setCode(e.target.value.replace(/\D/g, "").slice(0, 6))
                  }
                  placeholder="123456"
                  required
                  maxLength={6}
                  className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 text-center text-lg font-mono tracking-widest"
                  disabled={loading}
                />
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                  Enter the 6-digit code from your email
                </p>
              </div>

              <Button
                type="submit"
                disabled={loading || code.length !== 6}
                className="w-full bg-green-600 hover:bg-green-700 text-white font-medium py-2 px-4 rounded-lg transition-colors"
              >
                {loading ? (
                  <div className="flex items-center justify-center">
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin mr-2"></div>
                    Verifying...
                  </div>
                ) : (
                  "Verify & Sign In"
                )}
              </Button>

              <div className="flex gap-2">
                <Button
                  type="button"
                  onClick={isReturningUser ? goBackToEmail : goBackToDetails}
                  variant="outline"
                  className="flex-1 border-slate-300 text-slate-700 hover:bg-slate-50"
                  disabled={loading}
                >
                  Back
                </Button>
                <Button
                  type="button"
                  onClick={resendOTP}
                  variant="outline"
                  className="flex-1 border-blue-300 text-blue-700 hover:bg-blue-50"
                  disabled={loading}
                >
                  Resend Code
                </Button>
              </div>
            </form>
          )}

          <div className="mt-6 text-center">
            <p className="text-xs text-slate-500 dark:text-slate-400">
              🔒 Passwordless authentication with one-time codes
            </p>
          </div>
        </CardContent>
      </Card>
    </div>
  );
};

export default AuthPage;
