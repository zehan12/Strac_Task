require("dotenv").config();
const express = require("express");
const session = require("express-session");
const passport = require("passport");
const OIDCStrategy = require("passport-azure-ad").OIDCStrategy;
const axios = require("axios");

const app = express();

// Azure AD Configuration
const config = {
    identityMetadata: `https://login.microsoftonline.com/${process.env.TENANT_ID}/v2.0/.well-known/openid-configuration`,
    clientID: process.env.CLIENT_ID,
    responseType: "code",
    responseMode: "query",
    redirectUrl: "http://localhost:3000/auth/callback",
    allowHttpForRedirectUrl: true,
    clientSecret: process.env.CLIENT_SECRET,
    validateIssuer: false,
    passReqToCallback: false,
    scope: [
        "openid",
        "profile",
        "offline_access",
        "User.Read",
        "Files.ReadWrite",
    ],
};

// Configure session
app.use(session({ secret: "secret", resave: false, saveUninitialized: false }));

// Configure passport
passport.use(
    new OIDCStrategy(
        config,
        (iss, sub, profile, accessToken, refreshToken, params, done) => {
            process.nextTick(() => {
                profile.accessToken = accessToken;
                profile.refreshToken = refreshToken;
                return done(null, profile);
            });
        }
    )
);

passport.serializeUser((user, done) => done(null, user));
passport.deserializeUser((obj, done) => done(null, obj));

app.use(passport.initialize());
app.use(passport.session());

// Routes
app.get(
    "/login",
    passport.authenticate("azuread-openidconnect", { failureRedirect: "/" }),
    (req, res) => {
        res.redirect("/");
    }
);

app.get(
    "/auth/callback",
    passport.authenticate("azuread-openidconnect", { failureRedirect: "/" }),
    (req, res) => {
        res.redirect("/");
    }
);

app.get("/logout", (req, res) => {
    req.session.destroy((err) => {
        req.logout();
        res.redirect("/");
    });
});

app.get("/", (req, res) => {
    res.send(req.isAuthenticated() ? "Logged in" : "Logged out");
});

app.get("/files", async (req, res) => {
    if (!req.isAuthenticated()) {
        return res.redirect("/login");
    }

    const accessToken = req.user.accessToken;

    try {
        const response = await axios.get(
            "https://graph.microsoft.com/v1.0/me/drive/root/children",
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                },
            }
        );

        console.log("Response:", response.data);
        res.send(response.data);
    } catch (error) {
        if (error.response) {
            console.error("Response data:", error.response.data);
            console.error("Response status:", error.response.status);
            console.error("Response headers:", error.response.headers);
        } else if (error.request) {
            console.error("Request error:", error.request);
        } else {
            console.error("Error:", error.message);
        }
        res.status(500).send("Error retrieving files");
    }
});

module.exports = app;
